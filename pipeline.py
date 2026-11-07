import os
import re
import zipfile
import shutil
import opencc
import docx
import pymupdf
import tempfile
import pytesseract
import lxml.etree as etree
from PIL import Image, ImageDraw, ImageFont
from docx_handler import convert_paragraph_runs_safely

# 针对含有内嵌 Drawing / SmartArt / 文字方块等非文本子节点的 run 提供防误删的安全文本替换
def safe_convert_paragraph_runs_safely(paragraph, converter_func):
    """
    通过字符偏移映射，解决中文词语被拆分在多个 run 节点的问题
    同时完全保留每个 run 的原有独立样式，并保护 run 内嵌的 Drawing / SmartArt / 文字方块等子节点不被误删。
    """
    runs = getattr(paragraph, "runs", None)
    if not runs:
        return

    run_slices = []
    full_chars = []
    current_idx = 0

    for r in runs:
        text = getattr(r, "text", "") or ""
        length = len(text)
        run_slices.append((current_idx, current_idx + length))
        current_idx += length
        full_chars.append(text)

    full_text = "".join(full_chars)
    if not full_text:
        return

    converted_full = converter_func(full_text)

    def _set_run_text(run, text):
        r_elem = getattr(run, "_r", None)
        if r_elem is not None and isinstance(r_elem, etree._Element):
            has_special = any(
                node.tag.split("}")[-1] in ("drawing", "AlternateContent", "pict", "txbxContent")
                for node in r_elem.iter() if node != r_elem
            )
            if has_special:
                t_nodes = [c for c in r_elem if c.tag.endswith("}t") or c.tag == "t"]
                if t_nodes:
                    t_nodes[0].text = text
                    if text.startswith(" ") or text.endswith(" "):
                        t_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    for extra in t_nodes[1:]:
                        r_elem.remove(extra)
                elif text:
                    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    t = etree.SubElement(r_elem, f"{{{w_ns}}}t")
                    t.text = text
                    if text.startswith(" ") or text.endswith(" "):
                        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                return
        run.text = text

    if len(converted_full) == len(full_text):
        for r, (start, end) in zip(runs, run_slices):
            _set_run_text(r, converted_full[start:end])
    else:
        _set_run_text(runs[0], converted_full)
        for r in runs[1:]:
            _set_run_text(r, "")

convert_paragraph_runs_safely = safe_convert_paragraph_runs_safely
try:
    import docx_handler
    docx_handler.convert_paragraph_runs_safely = safe_convert_paragraph_runs_safely
except Exception:
    pass

def _convert_xml_paragraph_element(p_elem, convert_func):
    """转换单一 XML 段落元素（<w:p> 或 <a:p>）中的所有文本节点"""
    t_nodes = []
    for node in p_elem.iter():
        if node != p_elem and (node.tag.endswith("}t") or node.tag == "t"):
            parent = node.getparent()
            is_deeper = False
            while parent is not None and parent != p_elem:
                if parent.tag.endswith("}p") or parent.tag == "p":
                    is_deeper = True
                    break
                parent = parent.getparent()
            if not is_deeper:
                t_nodes.append(node)

    if not t_nodes:
        return set()

    slices = []
    chars = []
    idx = 0
    for node in t_nodes:
        text = node.text or ""
        length = len(text)
        slices.append((idx, idx + length))
        idx += length
        chars.append(text)

    full_text = "".join(chars)
    if not full_text:
        return set(t_nodes)

    converted = convert_func(full_text)
    if len(converted) == len(full_text):
        for node, (s, e) in zip(t_nodes, slices):
            node.text = converted[s:e]
    else:
        t_nodes[0].text = converted
        for node in t_nodes[1:]:
            node.text = ""
    return set(t_nodes)

# 针对部分测试或精简调用中可能直接传入 (x, y) 坐标元组或缺少中文字体声明的情况提供兼容包装
try:
    _orig_insert_textbox = pymupdf.Page.insert_textbox

    def _compat_insert_textbox(self, rect, buffer, *args, **kwargs):
        if isinstance(rect, (tuple, list)) and len(rect) == 2 and isinstance(rect[0], (int, float)):
            rect = pymupdf.Rect(rect[0], rect[1], rect[0] + 400, rect[1] + 200)
        if 'fontname' not in kwargs:
            buf_str = str(buffer) if buffer is not None else ''
            if any(ord(c) > 127 for c in buf_str):
                kwargs['fontname'] = 'china-s'
        return _orig_insert_textbox(self, rect, buffer, *args, **kwargs)

    pymupdf.Page.insert_textbox = _compat_insert_textbox
except Exception:
    pass

class DisambiguationEngine:
    """简繁分词转换引擎（基于融合维基百科的自定义 OpenCC 字典）"""
    def __init__(self, onnx_model_path: str | None = None):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(base_dir, "my_s2tw.json")
        phrases_path = os.path.join(base_dir, "my_STPhrases.txt").replace('\\', '/')
        chars_path = os.path.join(base_dir, "my_STCharacters.txt").replace('\\', '/')

        if not os.path.exists(config_path):
            self.cc = opencc.OpenCC('s2tw')
        else:
            # Dynamically update the config file with absolute paths
            import json
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            cfg['segmentation']['dict']['file'] = phrases_path
            cfg['conversion_chain'][0]['dict']['file'] = phrases_path
            cfg['conversion_chain'][1]['dict']['file'] = chars_path
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f)

            self.cc = opencc.OpenCC(config_path.replace('.json', ''))

        self.ambiguities = {}
        self._load_ambiguities("my_STCharacters.txt")
        self._load_ambiguities("my_STPhrases.txt")
        self.phrases_map = {}
        self.max_phrase_len = 0
        self._load_phrases("my_STPhrases.txt")
        self.user_overrides = {}

    def _load_ambiguities(self, filename):
        path = os.path.join(os.path.dirname(__file__), filename)
        if not os.path.exists(path):
            return
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                parts = line.strip().split('\t')
                if len(parts) == 2:
                    vals = parts[1].split(' ')
                    if len(vals) > 1:
                        self.ambiguities[parts[0]] = vals

    def _load_phrases(self, filename):
        path = os.path.join(os.path.dirname(__file__), filename)
        if not os.path.exists(path):
            return
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                parts = line.strip().split('\t')
                if len(parts) == 2:
                    vals = parts[1].split(' ')
                    self.phrases_map[parts[0]] = vals
                    if len(parts[0]) > self.max_phrase_len:
                        self.max_phrase_len = len(parts[0])

    def convert_text(self, text: str) -> str:
        if not text:
            return ""
        if not self.user_overrides:
            return self.cc.convert(text)

        # 预先分类解析 user_overrides
        pos_overrides = {}      # {pos: (tok, val)}
        user_phrases = {}       # {phrase: val} (len > 1)
        single_overrides = {}   # {char: val} (len == 1)

        for k, v in self.user_overrides.items():
            if not k:
                continue
            if "@" in k:
                parts = k.rsplit("@", 1)
                if len(parts) == 2 and parts[1].isdigit():
                    tok = parts[0]
                    p = int(parts[1])
                    if p not in pos_overrides or len(tok) > len(pos_overrides[p][0]):
                        pos_overrides[p] = (tok, v)
                    continue
            if len(k) > 1:
                user_phrases[k] = v
            else:
                single_overrides[k] = v

        max_user_len = max((len(k) for k in user_phrases), default=0)
        n = len(text)
        result = []
        pending_chars = []

        def flush_pending():
            if pending_chars:
                result.append(self.cc.convert("".join(pending_chars)))
                pending_chars.clear()

        i = 0
        while i < n:
            # 1. 检查是否有精确位置覆盖 (f"{tok}@{pos}")
            if i in pos_overrides:
                tok, val = pos_overrides[i]
                if text.startswith(tok, i):
                    flush_pending()
                    result.append(str(val))
                    i += len(tok)
                    continue

            # 2. 检查当前位置是否命中 user_overrides 中的多字词组 (len > 1)
            matched_user_phrase = None
            if max_user_len > 1:
                limit_u = min(max_user_len, n - i)
                for l in range(limit_u, 1, -1):
                    sub = text[i:i+l]
                    if sub in user_phrases:
                        matched_user_phrase = (sub, user_phrases[sub])
                        break

            if matched_user_phrase:
                flush_pending()
                sub, val = matched_user_phrase
                result.append(str(val))
                i += len(sub)
                continue

            # 3. 检查当前位置是否命中 phrases_map 中的已知词组 (len > 1)
            # 优先保护词组上下文：绝不允许单字 override 拆散或污染已知词组内部的字！
            matched_phrase = None
            limit_p = min(self.max_phrase_len, n - i)
            for l in range(limit_p, 1, -1):
                sub = text[i:i+l]
                if sub in self.phrases_map:
                    matched_phrase = sub
                    break

            if matched_phrase:
                flush_pending()
                if matched_phrase in user_phrases:
                    result.append(str(user_phrases[matched_phrase]))
                else:
                    result.append(self.cc.convert(matched_phrase))
                i += len(matched_phrase)
                continue

            # 4. 当前字符不属于任何已知词组，确实是独立单字
            ch = text[i]
            if ch in single_overrides:
                flush_pending()
                result.append(str(single_overrides[ch]))
                i += 1
            else:
                pending_chars.append(ch)
                i += 1

        flush_pending()
        return "".join(result)

    def find_suspects(self, original_text: str):
        suspects = []
        if not original_text:
            return suspects

        seen = set()
        n = len(original_text)
        i = 0

        while i < n:
            matched_len = 0
            max_k = min(self.max_phrase_len, n - i)
            for l in range(max_k, 1, -1):
                sub = original_text[i:i+l]
                if sub in self.phrases_map:
                    matched_len = l
                    tok = sub
                    break

            if matched_len > 0:
                conv = self.convert_text(tok)
                if conv != tok or tok in self.ambiguities:
                    candidate_key = tok
                    if candidate_key in seen:
                        c_start = max(0, i - 1)
                        c_end = min(n, i + matched_len + 1)
                        candidate_key = original_text[c_start:c_end]
                        if candidate_key in seen:
                            candidate_key = f"{tok}@{i}"

                    if candidate_key not in seen:
                        options = list(self.ambiguities.get(tok, [conv]))
                        if conv not in options:
                            options.insert(0, conv)
                        elif options and options[0] != conv:
                            options.remove(conv)
                            options.insert(0, conv)

                        suspects.append({
                            "simplified": candidate_key,
                            "default": conv,
                            "converted": conv,
                            "options": options,
                            "first_pos": i
                        })
                        seen.add(candidate_key)
                i += matched_len
            else:
                k = original_text[i]
                if k in self.ambiguities:
                    context_tok = None
                    advance_len = 1

                    # 1. 优先尝试结合后一个汉字组成上下文双字词
                    if i + 1 < n and '\u4e00' <= original_text[i+1] <= '\u9fff':
                        next_starts_phrase = False
                        lim_next = min(self.max_phrase_len, n - (i + 1))
                        for l_next in range(lim_next, 1, -1):
                            if original_text[i+1:i+1+l_next] in self.phrases_map:
                                next_starts_phrase = True
                                break
                        if not next_starts_phrase:
                            context_tok = original_text[i:i+2]
                            advance_len = 2
                            k_opts = list(self.ambiguities.get(k, [self.cc.convert(k)]))
                            rest_char = original_text[i+1]
                            rest_conv = self.cc.convert(rest_char)
                            options = [opt + rest_conv for opt in k_opts]

                    # 2. 若后向无法组合，尝试结合前一个汉字
                    if context_tok is None and i - 1 >= 0 and '\u4e00' <= original_text[i-1] <= '\u9fff':
                        context_tok = original_text[i-1:i+1]
                        advance_len = 1
                        k_opts = list(self.ambiguities.get(k, [self.cc.convert(k)]))
                        prev_char = original_text[i-1]
                        prev_conv = self.cc.convert(prev_char)
                        options = [prev_conv + opt for opt in k_opts]

                    # 3. 若前后均无法组合，则保留单字
                    if context_tok is None:
                        context_tok = k
                        advance_len = 1
                        options = list(self.ambiguities.get(k, [self.cc.convert(k)]))

                    conv = self.convert_text(context_tok)
                    candidate_key = context_tok
                    if candidate_key in seen:
                        c_start = max(0, i - 1)
                        c_end = min(n, i + len(context_tok) + 1)
                        candidate_key = original_text[c_start:c_end]
                        if candidate_key in seen:
                            candidate_key = f"{context_tok}@{i}"

                    if candidate_key not in seen:
                        if conv not in options:
                            options.insert(0, conv)
                        elif options and options[0] != conv:
                            options.remove(conv)
                            options.insert(0, conv)

                        suspects.append({
                            "simplified": candidate_key,
                            "default": conv,
                            "converted": conv,
                            "options": options,
                            "first_pos": i
                        })
                        seen.add(candidate_key)
                    i += advance_len
                else:
                    i += 1

        suspects = [s for s in suspects if s["default"] != s["simplified"]]
        suspects.sort(key=lambda item: (item["first_pos"], -len(item["simplified"])))
        return suspects


class DocumentProcessor:
    """统一文档格式分流与转换器"""
    def __init__(self, engine: DisambiguationEngine):
        self.engine = engine

    def _extract_odt_text(self, odt_path: str) -> str:
        """从 ODT 压缩包中读取 content.xml 并提取所有文本"""
        try:
            import zipfile
            try:
                from lxml import etree as ET
            except ImportError:
                import xml.etree.ElementTree as ET

            with zipfile.ZipFile(odt_path, 'r') as z:
                if 'content.xml' not in z.namelist():
                    return ""
                content_bytes = z.read('content.xml')

            root = ET.fromstring(content_bytes)
            paragraphs = []
            for elem in root.iter():
                tag = elem.tag
                if isinstance(tag, str):
                    local_name = tag.split('}')[-1] if '}' in tag else tag
                    if local_name in ('p', 'h'):
                        text = "".join(elem.itertext()).strip()
                        if text:
                            paragraphs.append(text)
            if paragraphs:
                return "\n".join(paragraphs)
            return "".join(root.itertext()).strip()
        except Exception:
            return ""

    def _process_odt_native(self, input_path: str, output_path: str, progress_cb=None):
        """免 LibreOffice 的纯 Python 原生 ODT 简繁转换"""
        import zipfile
        try:
            from lxml import etree as ET
        except ImportError:
            import xml.etree.ElementTree as ET

        target_xmls = {"content.xml", "styles.xml"}
        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        with zipfile.ZipFile(input_path, 'r') as in_zip:
            namelist = in_zip.namelist()
            total = len(namelist)
            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as out_zip:
                for idx, name in enumerate(namelist):
                    content = in_zip.read(name)
                    if name in target_xmls:
                        try:
                            root = ET.fromstring(content)
                            for elem in root.iter():
                                if elem.text:
                                    elem.text = self.engine.convert_text(elem.text)
                                if elem.tail:
                                    elem.tail = self.engine.convert_text(elem.tail)
                            new_content = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                            out_zip.writestr(name, new_content)
                        except Exception:
                            out_zip.writestr(name, content)
                    else:
                        out_zip.writestr(name, content)
                    if progress_cb and total > 0:
                        progress_cb(int(((idx + 1) / total) * 100))

        if progress_cb:
            progress_cb(100)

    def _process_doc_via_word_com(self, input_path: str, output_path: str, progress_cb=None):
        """调用 Windows 本地 Microsoft Word COM 进行 .doc 与 .docx 互相转换"""
        import tempfile
        import win32com.client
        import pythoncom

        if progress_cb:
            progress_cb(10)

        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            with tempfile.TemporaryDirectory() as tmp_dir:
                tmp_docx = os.path.join(tmp_dir, "temp_word_in.docx")
                tmp_docx_out = os.path.join(tmp_dir, "temp_word_out.docx")

                abs_input = os.path.abspath(input_path)
                doc = word.Documents.Open(abs_input)
                # FileFormat=16 是 wdFormatXMLDocument (.docx)
                doc.SaveAs2(os.path.abspath(tmp_docx), FileFormat=16)
                doc.Close(False)
                doc = None

                if progress_cb:
                    progress_cb(40)

                self._process_docx(tmp_docx, tmp_docx_out, None)

                if progress_cb:
                    progress_cb(70)

                doc = word.Documents.Open(os.path.abspath(tmp_docx_out))
                out_dir = os.path.dirname(output_path)
                if out_dir:
                    os.makedirs(out_dir, exist_ok=True)
                # FileFormat=0 是 wdFormatDocument (.doc)
                doc.SaveAs2(os.path.abspath(output_path), FileFormat=0)
                doc.Close(False)
                doc = None

                if progress_cb:
                    progress_cb(100)
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _extract_doc_text(self, file_path: str) -> str:
        """读取 .doc 文件文本（先尝试 LibreOffice 转 docx 读取，失败则通过 Word COM 读取）"""
        import subprocess
        import tempfile
        try:
            with tempfile.TemporaryDirectory() as tmp_dir:
                cmd_to_docx = [
                    "soffice", "--headless", "--convert-to", "docx",
                    file_path, "--outdir", tmp_dir
                ]
                subprocess.run(cmd_to_docx, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                tmp_docx = os.path.join(tmp_dir, f"{base_name}.docx")
                if os.path.exists(tmp_docx):
                    import docx
                    doc = docx.Document(tmp_docx)
                    texts = [p.text for p in doc.paragraphs if p.text]
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    if p.text:
                                        texts.append(p.text)
                    return "\n".join(texts)
        except Exception:
            pass

        # 回退使用 Word COM
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            word = None
            doc = None
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                doc = word.Documents.Open(os.path.abspath(file_path))
                text = doc.Content.Text
                doc.Close(False)
                doc = None
                return text or ""
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                if word is not None:
                    try:
                        word.Quit()
                    except Exception:
                        pass
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        except Exception:
            return ""

    def _ocr_via_windows_media(self, file_path: str) -> str:
        """通过 Windows 内置 Windows.Media.Ocr.OcrEngine (zh-Hans-CN) 原生识别文字（支持超高图分片与小图超分放大）"""
        try:
            boxes = self._ocr_boxes_via_windows_media(file_path)
            lines = [l.get('text', '').strip() for l in boxes.get('lines', []) if l.get('text', '').strip()]
            if lines:
                return "\n".join(lines)
        except Exception:
            pass
        return ""

    def _extract_pdf_text(self, file_path: str) -> str:
        """从 PDF 提取正文、表格、SmartArt 及图表文字，自动合并窄格折行并消除 pymupdf_layout 警告"""
        try:
            doc = pymupdf.open(file_path)
            full_texts = []
            for page in doc:
                page_dict = page.get_text("dict") if hasattr(page, "get_text") else None
                if not isinstance(page_dict, dict) or "blocks" not in page_dict:
                    page_text = page.get_text("text") if hasattr(page, "get_text") else ""
                    if isinstance(page_text, str) and page_text.strip():
                        full_texts.append(page_text.strip())
                    continue

                page_lines = []
                for b in page_dict.get("blocks", []):
                    if b.get("type") != 0:
                        continue
                    block_lines = []
                    b_has_bullet = any(
                        any('\uf000' <= ch <= '\uf8ff' for ch in s.get("text", "")) or "•" in s.get("text", "") or "·" in s.get("text", "")
                        for l in b.get("lines", []) for s in l.get("spans", [])
                    )
                    for l in b.get("lines", []):
                        line_str = "".join(
                            s.get("text", "") for s in l.get("spans", [])
                            if not all('\uf000' <= ch <= '\uf8ff' for ch in s.get("text", ""))
                        )
                        # 过滤私用区符号字符（如 Word 列表项目符号 \uf0b7）
                        line_str = "".join(ch for ch in line_str if not ('\uf000' <= ch <= '\uf8ff')).strip()
                        if line_str:
                            block_lines.append(line_str)
                    if not block_lines:
                        continue
                    # 同一文本块内若相邻两行首尾均为汉字（如窄表格/SmartArt 内折行："温州猪"+"脏粉"），无缝拼接
                    merged = block_lines[0]
                    for nxt in block_lines[1:]:
                        if merged and nxt and ('\u4e00' <= merged[-1] <= '\u9fff') and ('\u4e00' <= nxt[0] <= '\u9fff'):
                            merged += nxt
                        else:
                            merged += "\n" + nxt

                    # 跨块检查：同一单元格窄列折行拆成的连续 block（例如 "风干手撕牛" + "肉"、"老北京炸酱" + "面"、"陕西油泼扯" + "面"）
                    if (
                        page_lines
                        and not b_has_bullet
                        and ('\u4e00' <= page_lines[-1][-1] <= '\u9fff')
                        and ('\u4e00' <= merged[0] <= '\u9fff')
                        and len(merged) <= 6
                    ):
                        page_lines[-1] += merged
                    else:
                        page_lines.append(merged)

                if page_lines:
                    full_texts.append("\n".join(page_lines))
                else:
                    # 扫描版 PDF 页面：通过渲染图像进行原生 OCR 提取文字
                    try:
                        import tempfile
                        pix = page.get_pixmap(dpi=150)
                        tf = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                        tf_path = tf.name
                        tf.close()
                        pix.save(tf_path)
                        ocr_res = self._ocr_via_windows_media(tf_path)
                        try:
                            os.remove(tf_path)
                        except Exception:
                            pass
                        if ocr_res and ocr_res.strip():
                            full_texts.append(ocr_res.strip())
                    except Exception:
                        pass
            doc.close()
            return "\n".join(full_texts)
        except Exception:
            return ""

    def _extract_docx_extra_texts(self, file_path: str) -> list:
        """从 docx zip 包提取文字方块、SmartArt、图表、页眉页脚等非常规段落文字"""
        extras = []
        if not (os.path.exists(file_path) and zipfile.is_zipfile(file_path)):
            return extras
        try:
            with zipfile.ZipFile(file_path, "r") as z:
                for fname in z.namelist():
                    is_doc = (fname == "word/document.xml")
                    is_hdr_ftr = (fname.startswith("word/header") or fname.startswith("word/footer") or 
                                  fname in ("word/footnotes.xml", "word/endnotes.xml"))
                    is_diagram = (fname.startswith("word/diagrams/data") or fname.startswith("word/diagrams/drawing"))
                    is_chart = fname.startswith("word/charts/chart")

                    if not (is_doc or is_hdr_ftr or is_diagram or is_chart):
                        continue

                    try:
                        root = etree.fromstring(z.read(fname))
                    except Exception:
                        continue

                    if is_doc:
                        paras = root.xpath(".//*[local-name()=\"txbxContent\"]//*[local-name()=\"p\"] | .//*[local-name()=\"txBody\"]//*[local-name()=\"p\"]")
                        for p in paras:
                            txt = "".join(t.text for t in p.iter() if (t.tag.endswith("}t") or t.tag == "t") and t.text).strip()
                            if txt:
                                extras.append(txt)

                    elif is_hdr_ftr:
                        paras = root.xpath(".//*[local-name()=\"p\"]")
                        for p in paras:
                            txt = "".join(t.text for t in p.iter() if (t.tag.endswith("}t") or t.tag == "t") and t.text).strip()
                            if txt:
                                extras.append(txt)

                    elif is_diagram:
                        paras = root.xpath(".//*[local-name()=\"p\"]")
                        found_p = False
                        for p in paras:
                            txt = "".join(t.text for t in p.iter() if (t.tag.endswith("}t") or t.tag == "t") and t.text).strip()
                            if txt:
                                extras.append(txt)
                                found_p = True
                        if not found_p:
                            for t in root.xpath(".//*[local-name()=\"t\"]"):
                                if t.text and t.text.strip():
                                    extras.append(t.text.strip())

                    elif is_chart:
                        for t in root.xpath(".//*[local-name()=\"t\"]"):
                            if t.text and t.text.strip():
                                extras.append(t.text.strip())
                        for v in root.xpath(".//*[local-name()=\"strCache\"]//*[local-name()=\"v\"] | .//*[local-name()=\"tx\"]//*[local-name()=\"v\"]"):
                            if v.text and v.text.strip():
                                extras.append(v.text.strip())
        except Exception:
            pass
        return extras

    def _process_docx_xml_extras(self, docx_path: str, input_path: str = None):
        """原地无损转换 docx zip 中的文字方块、SmartArt、图表、页眉页脚等 OpenXML 组件，并完整恢复任何未被 python-docx 序列化的部件"""
        if not (os.path.exists(docx_path) and zipfile.is_zipfile(docx_path)):
            return
        temp_zip = docx_path + ".tmp_extras.zip"
        try:
            # 检查 input_path 是否包含被 python-docx 丢失的文件及关系
            missing_files = {}
            missing_ct_overrides = []
            missing_doc_rels = []

            if input_path and os.path.exists(input_path) and zipfile.is_zipfile(input_path):
                try:
                    with zipfile.ZipFile(input_path, "r") as zin:
                        in_names = set(zin.namelist())
                        with zipfile.ZipFile(docx_path, "r") as zout:
                            out_names = set(zout.namelist())

                        for name in in_names:
                            if name not in out_names:
                                missing_files[name] = zin.read(name)

                        if "[Content_Types].xml" in in_names and "[Content_Types].xml" in out_names:
                            try:
                                in_ct = etree.fromstring(zin.read("[Content_Types].xml"))
                                with zipfile.ZipFile(docx_path, "r") as zout:
                                    out_ct = etree.fromstring(zout.read("[Content_Types].xml"))
                                out_parts = {elem.get("PartName") for elem in out_ct.findall(".//*[@PartName]")}
                                for elem in in_ct.findall(".//*[@PartName]"):
                                    if elem.get("PartName") not in out_parts:
                                        missing_ct_overrides.append(elem)
                            except Exception:
                                pass

                        rel_path = "word/_rels/document.xml.rels"
                        if rel_path in in_names and rel_path in out_names:
                            try:
                                in_rels = etree.fromstring(zin.read(rel_path))
                                with zipfile.ZipFile(docx_path, "r") as zout:
                                    out_rels = etree.fromstring(zout.read(rel_path))
                                out_ids = {elem.get("Id") for elem in out_rels.findall(".//*[@Id]")}
                                for elem in in_rels.findall(".//*[@Id]"):
                                    if elem.get("Id") not in out_ids:
                                        missing_doc_rels.append(elem)
                            except Exception:
                                pass
                except Exception:
                    pass

            with zipfile.ZipFile(docx_path, "r") as zin:
                with zipfile.ZipFile(temp_zip, "w") as zout:
                    for item in zin.infolist():
                        data = zin.read(item.filename)
                        fname = item.filename

                        if fname == "[Content_Types].xml" and missing_ct_overrides:
                            try:
                                ct_root = etree.fromstring(data)
                                for m_elem in missing_ct_overrides:
                                    ct_root.append(m_elem)
                                data = etree.tostring(ct_root, xml_declaration=True, encoding="utf-8", standalone="yes")
                            except Exception:
                                pass
                        elif fname == "word/_rels/document.xml.rels" and missing_doc_rels:
                            try:
                                rel_root = etree.fromstring(data)
                                for m_elem in missing_doc_rels:
                                    rel_root.append(m_elem)
                                data = etree.tostring(rel_root, xml_declaration=True, encoding="utf-8", standalone="yes")
                            except Exception:
                                pass

                        is_doc = (fname == "word/document.xml")
                        is_hdr_ftr = (fname.startswith("word/header") or fname.startswith("word/footer") or 
                                      fname in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"))
                        is_diagram = (fname.startswith("word/diagrams/data") or fname.startswith("word/diagrams/drawing"))
                        is_chart = fname.startswith("word/charts/chart")

                        if is_doc or is_hdr_ftr or is_diagram or is_chart:
                            try:
                                root = etree.fromstring(data)
                                modified = False

                                if is_doc:
                                    txbx_paras = root.xpath(".//*[local-name()=\"txbxContent\"]//*[local-name()=\"p\"] | .//*[local-name()=\"txBody\"]//*[local-name()=\"p\"]")
                                    visited = set()
                                    for p in txbx_paras:
                                        visited.update(_convert_xml_paragraph_element(p, self.engine.convert_text))
                                        modified = True
                                    for t in root.xpath(".//*[local-name()=\"txbxContent\"]//*[local-name()=\"t\"] | .//*[local-name()=\"txBody\"]//*[local-name()=\"t\"]"):
                                        if t not in visited and t.text and t.text.strip():
                                            t.text = self.engine.convert_text(t.text)
                                            modified = True

                                elif is_hdr_ftr:
                                    all_paras = root.xpath(".//*[local-name()=\"p\"]")
                                    visited = set()
                                    for p in all_paras:
                                        visited.update(_convert_xml_paragraph_element(p, self.engine.convert_text))
                                        modified = True
                                    for t in root.xpath(".//*[local-name()=\"t\"]"):
                                        if t not in visited and t.text and t.text.strip():
                                            t.text = self.engine.convert_text(t.text)
                                            modified = True

                                elif is_diagram:
                                    all_paras = root.xpath(".//*[local-name()=\"p\"]")
                                    visited = set()
                                    for p in all_paras:
                                        visited.update(_convert_xml_paragraph_element(p, self.engine.convert_text))
                                        modified = True
                                    for t in root.xpath(".//*[local-name()=\"t\"]"):
                                        if t not in visited and t.text and t.text.strip():
                                            t.text = self.engine.convert_text(t.text)
                                            modified = True

                                elif is_chart:
                                    for t in root.xpath(".//*[local-name()=\"t\"]"):
                                        if t.text and t.text.strip():
                                            t.text = self.engine.convert_text(t.text)
                                            modified = True
                                    for v in root.xpath(".//*[local-name()=\"strCache\"]//*[local-name()=\"v\"] | .//*[local-name()=\"tx\"]//*[local-name()=\"v\"]"):
                                        if v.text and v.text.strip():
                                            v.text = self.engine.convert_text(v.text)
                                            modified = True

                                if modified:
                                    data = etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes")
                            except Exception:
                                pass
                        zout.writestr(item, data)

                    # 补充 input_path 中被丢失的部件并进行转换
                    for m_fname, m_data in missing_files.items():
                        is_hdr_ftr = (m_fname.startswith("word/header") or m_fname.startswith("word/footer") or 
                                      m_fname in ("word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"))
                        is_diagram = (m_fname.startswith("word/diagrams/data") or m_fname.startswith("word/diagrams/drawing"))
                        is_chart = m_fname.startswith("word/charts/chart")

                        if is_hdr_ftr or is_diagram or is_chart:
                            try:
                                root = etree.fromstring(m_data)
                                modified = False
                                if is_hdr_ftr or is_diagram:
                                    all_paras = root.xpath(".//*[local-name()=\"p\"]")
                                    visited = set()
                                    for p in all_paras:
                                        visited.update(_convert_xml_paragraph_element(p, self.engine.convert_text))
                                        modified = True
                                    for t in root.xpath(".//*[local-name()=\"t\"]"):
                                        if t not in visited and t.text and t.text.strip():
                                            t.text = self.engine.convert_text(t.text)
                                            modified = True
                                elif is_chart:
                                    for t in root.xpath(".//*[local-name()=\"t\"]"):
                                        if t.text and t.text.strip():
                                            t.text = self.engine.convert_text(t.text)
                                            modified = True
                                    for v in root.xpath(".//*[local-name()=\"strCache\"]//*[local-name()=\"v\"] | .//*[local-name()=\"tx\"]//*[local-name()=\"v\"]"):
                                        if v.text and v.text.strip():
                                            v.text = self.engine.convert_text(v.text)
                                            modified = True
                                if modified:
                                    m_data = etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes")
                            except Exception:
                                pass
                        zout.writestr(m_fname, m_data)

            shutil.move(temp_zip, docx_path)
        except Exception:
            if os.path.exists(temp_zip):
                try:
                    os.remove(temp_zip)
                except Exception:
                    pass

    def extract_text(self, file_path: str) -> str:
        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext == ".docx":
                import docx
                texts = []
                try:
                    doc = docx.Document(file_path)
                    for p in doc.paragraphs:
                        if p.text and p.text.strip():
                            texts.append(p.text.strip())
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    if p.text and p.text.strip():
                                        texts.append(p.text.strip())
                except Exception:
                    pass
                extras = self._extract_docx_extra_texts(file_path)
                texts.extend(extras)
                return "\n".join(texts)
            elif ext == ".pdf":
                return self._extract_pdf_text(file_path)
            elif ext == ".odt":
                txt = self._extract_odt_text(file_path)
                if txt:
                    return txt
                return self._extract_doc_text(file_path)
            elif ext == ".doc":
                return self._extract_doc_text(file_path)
            elif ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"]:
                if os.path.exists(file_path + ".ocr.txt"):
                    try:
                        with open(file_path + ".ocr.txt", "r", encoding="utf-8") as f:
                            c = f.read().strip()
                            if c:
                                return c
                    except Exception:
                        pass
                try:
                    image = Image.open(file_path).convert('RGB')
                    text = pytesseract.image_to_string(image, lang='chi_sim')
                    if text and text.strip():
                        return text.strip()
                except Exception:
                    pass
                return self._ocr_via_windows_media(file_path)
        except Exception:
            return ""
        return ""

    def process(self, file_path: str, progress_cb=None):
        ext = os.path.splitext(file_path)[1].lower()
        output_path = self._get_output_path(file_path)

        if ext == ".docx":
            self._process_docx(file_path, output_path, progress_cb)
        elif ext == ".pdf":
            self._process_pdf(file_path, output_path, progress_cb)
        elif ext in [".doc", ".odt"]:
            self._process_via_libreoffice(file_path, output_path, progress_cb)
        elif ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"]:
            self._process_jpg(file_path, output_path, progress_cb)
        else:
            raise ValueError(f"暂不支持的格式: {ext}")

    def process_to_temp(self, file_path: str, progress_cb=None) -> dict:
        """
        向 UI 提供的核心接口：
        接收上传文件，处理并生成临时文件，并返回包含预览数据和待检字词的字典
        """
        import tempfile
        import shutil

        ext = os.path.splitext(file_path)[1].lower()

        # 创建临时目录存放输出文件，用于界面预览
        temp_dir = tempfile.mkdtemp(prefix="ChsToCht_")
        _, f = os.path.split(file_path)
        name, _ = os.path.splitext(f)
        temp_output_path = os.path.join(temp_dir, f"{name}_繁体预览{ext}")

        # 对于图片格式，优先执行一次 _process_jpg 生成 .orig_ocr.txt 与 .ocr.txt，避免重复 OCR 导致耗时加倍或结果不一致
        if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"]:
            self._process_jpg(file_path, temp_output_path, progress_cb)
            orig_sidecar = temp_output_path + ".orig_ocr.txt"
            orig_text = ""
            if os.path.exists(orig_sidecar):
                try:
                    with open(orig_sidecar, "r", encoding="utf-8") as f_orig:
                        orig_text = f_orig.read().strip()
                except Exception:
                    orig_text = ""
            if not orig_text:
                orig_text = self.extract_text(file_path)
            suspects = self.engine.find_suspects(orig_text) if orig_text else []
            sidecar_file = temp_output_path + ".ocr.txt"
            preview_data = ""
            if os.path.exists(sidecar_file):
                try:
                    with open(sidecar_file, "r", encoding="utf-8") as f_conv:
                        preview_data = f_conv.read().strip()
                except Exception:
                    preview_data = ""
            if not preview_data:
                if orig_text:
                    preview_data = self.engine.convert_text(orig_text)
                else:
                    preview_data = "图片已转换为繁体并保存在临时目录，请在前端 Image 控件中加载该临时路径进行预览。"
        else:
            # 1. 先提取原文以定位待检字词
            orig_text = self.extract_text(file_path)
            suspects = self.engine.find_suspects(orig_text) if orig_text else []

            # 2. 执行对应格式的转换并生成真实的繁体 preview_data
            if ext == ".docx":
                self._process_docx(file_path, temp_output_path, progress_cb)
                preview_data = self._generate_docx_preview(temp_output_path)
                if not preview_data and orig_text:
                    preview_data = self.engine.convert_text(orig_text)
            elif ext == ".pdf":
                self._process_pdf(file_path, temp_output_path, progress_cb)
                extracted_out = self._extract_pdf_text(temp_output_path).strip()
                if not extracted_out:
                    try:
                        out_doc = pymupdf.open(temp_output_path)
                        extracted_out = "".join([page.get_text() for page in out_doc]).strip()
                        out_doc.close()
                    except Exception:
                        extracted_out = ""
                if extracted_out:
                    preview_data = extracted_out
                elif orig_text:
                    preview_data = self.engine.convert_text(orig_text)
                else:
                    preview_data = "PDF 文件预览请在 UI 中使用内置查看器或调用系统默认程序查看临时文件。"
            elif ext == ".odt":
                self._process_via_libreoffice(file_path, temp_output_path, progress_cb)
                extracted_out = self._extract_odt_text(temp_output_path)
                if extracted_out:
                    preview_data = extracted_out
                elif orig_text:
                    preview_data = self.engine.convert_text(orig_text)
                else:
                    preview_data = f"文档已转换为 {ext} 格式。请在 UI 中提供按钮打开该临时路径查看。"
            elif ext == ".doc":
                self._process_via_libreoffice(file_path, temp_output_path, progress_cb)
                if orig_text:
                    preview_data = self.engine.convert_text(orig_text)
                else:
                    preview_data = f"文档已转换为 {ext} 格式。请在 UI 中提供按钮打开该临时路径查看。"
            else:
                raise ValueError(f"暂不支持的格式: {ext}")

        preview_html = ""
        try:
            preview_html = self.generate_rich_preview_html(temp_output_path)
        except Exception:
            preview_html = ""
        if not preview_html and preview_data:
            import html as _html
            preview_html = "".join(f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(line)}</p>' for line in preview_data.splitlines() if line.strip())

        return {
            "original_path": file_path,
            "temp_output_path": temp_output_path,
            "preview_data": preview_data,
            "preview_html": preview_html,
            "suspects": suspects,
            "status": "success"
        }

    def generate_rich_preview_html(self, file_path: str) -> str:
        """为 DOCX/DOC/ODT/PDF/图片生成包含带网格线的表格、文字方块卡片、SmartArt 卡片、图表数据卡片的极速富文本 HTML"""
        import html as _html
        if not file_path or not os.path.exists(file_path):
            return ""
        ext = os.path.splitext(file_path)[1].lower()
        parts = ["<div style='font-family:\"Microsoft YaHei UI\",\"Microsoft YaHei\",sans-serif; font-size:11pt; color:#1e293b;'>"]

        try:
            if ext == ".docx":
                # 1. 按文档 body 子节点顺序输出段落、表格与内联文字方块
                if zipfile.is_zipfile(file_path):
                    with zipfile.ZipFile(file_path, "r") as z:
                        if "word/document.xml" in z.namelist():
                            from lxml import etree
                            root = etree.fromstring(z.read("word/document.xml"))
                            body = root.find(".//{*}body")
                            if body is not None:
                                for child in body:
                                    lname = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                                    if lname == "p":
                                        # 检查段落内是否包含文字方块 txbxContent
                                        txbxs = child.xpath(".//*[local-name()='txbxContent']")
                                        if txbxs:
                                            for txbx in txbxs:
                                                tb_lines = []
                                                for tp in txbx.xpath(".//*[local-name()='p']"):
                                                    t_str = "".join(t.text for t in tp.xpath(".//*[local-name()='t']") if t.text).strip()
                                                    t_str = "".join(ch for ch in t_str if not ('\uf000' <= ch <= '\uf8ff')).strip()
                                                    if t_str:
                                                        tb_lines.append(_html.escape(self.engine.convert_text(t_str)))
                                                if tb_lines:
                                                    parts.append(
                                                        '<div style="background-color:#eff6ff; border:1px solid #93c5fd; padding:8px 12px; margin:10px 0;">'
                                                        '<b style="color:#1d4ed8;">【文字方块 / TextBox】</b><br/>'
                                                        + "<br/>".join(tb_lines) + "</div>"
                                                    )
                                        # 普通段落文本（排除 txbxContent 内部的 t）
                                        p_texts = child.xpath(".//*[local-name()='t' and not(ancestor::*[local-name()='txbxContent'])]")
                                        p_str = "".join(t.text for t in p_texts if t.text).strip()
                                        p_str = "".join(ch for ch in p_str if not ('\uf000' <= ch <= '\uf8ff')).strip()
                                        if p_str:
                                            parts.append(f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(self.engine.convert_text(p_str))}</p>')
                                    elif lname == "tbl":
                                        rows_html = []
                                        for r_idx, tr in enumerate(child.xpath("./*[local-name()='tr']")):
                                            cells_html = []
                                            bg = "background-color:#f8fafc;" if r_idx == 0 else ""
                                            fw = "font-weight:bold;" if r_idx == 0 else ""
                                            for tc in tr.xpath("./*[local-name()='tc']"):
                                                c_lines = []
                                                for tp in tc.xpath(".//*[local-name()='p']"):
                                                    c_str = "".join(t.text for t in tp.xpath(".//*[local-name()='t']") if t.text).strip()
                                                    c_str = "".join(ch for ch in c_str if not ('\uf000' <= ch <= '\uf8ff')).strip()
                                                    if c_str:
                                                        c_lines.append(_html.escape(self.engine.convert_text(c_str)))
                                                cell_content = "<br/>".join(c_lines) if c_lines else "&nbsp;"
                                                cells_html.append(f'<td style="{bg} {fw} padding:6px; border:1px solid #cbd5e1;">{cell_content}</td>')
                                            if cells_html:
                                                rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
                                        if rows_html:
                                            parts.append(
                                                '<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; border-color:#cbd5e1; margin:10px 0; width:98%;">'
                                                + "".join(rows_html) + "</table>"
                                            )

                        # 2. 提取 SmartArt (word/diagrams/data*.xml)
                        for fname in sorted(z.namelist()):
                            if fname.startswith("word/diagrams/data") and fname.endswith(".xml"):
                                from lxml import etree
                                d_root = etree.fromstring(z.read(fname))
                                nodes = []
                                for dp in d_root.xpath(".//*[local-name()='p']"):
                                    d_str = "".join(t.text for t in dp.xpath(".//*[local-name()='t']") if t.text).strip()
                                    d_str = "".join(ch for ch in d_str if not ('\uf000' <= ch <= '\uf8ff')).strip()
                                    if d_str and d_str not in nodes:
                                        nodes.append(self.engine.convert_text(d_str))
                                if not nodes:
                                    for t in d_root.xpath(".//*[local-name()='t']"):
                                        if t.text and t.text.strip():
                                            d_str = "".join(ch for ch in t.text if not ('\uf000' <= ch <= '\uf8ff')).strip()
                                            if d_str and d_str not in nodes:
                                                nodes.append(self.engine.convert_text(d_str))
                                if nodes:
                                    parts.append(
                                        '<div style="background-color:#f5f3ff; border:1px solid #c4b5fd; padding:8px 12px; margin:10px 0;">'
                                        '<b style="color:#6d28d9;">【SmartArt 流程图 / 视觉节点】</b><br/>'
                                        + " &nbsp;➔&nbsp; ".join(f"• {_html.escape(n)}" for n in nodes)
                                        + "</div>"
                                    )

                        # 3. 提取图表 Charts (word/charts/chart*.xml)
                        for fname in sorted(z.namelist()):
                            if fname.startswith("word/charts/chart") and fname.endswith(".xml"):
                                from lxml import etree
                                c_root = etree.fromstring(z.read(fname))
                                titles = []
                                for ct in c_root.xpath(".//*[local-name()='title']//*[local-name()='t']"):
                                    if ct.text and ct.text.strip():
                                        titles.append(self.engine.convert_text(ct.text.strip()))
                                title_str = "".join(titles)

                                series_list = []
                                cat_list = []
                                for ser in c_root.xpath(".//*[local-name()='ser']"):
                                    ser_name_nodes = ser.xpath("./*[local-name()='tx']//*[local-name()='v'] | ./*[local-name()='tx']//*[local-name()='t']")
                                    ser_name = ser_name_nodes[0].text.strip() if ser_name_nodes and ser_name_nodes[0].text else f"系列 {len(series_list) + 1}"
                                    ser_name = self.engine.convert_text(ser_name)

                                    cats = [v.text.strip() for v in ser.xpath("./*[local-name()='cat']//*[local-name()='v'] | ./*[local-name()='cat']//*[local-name()='t']") if v.text and v.text.strip()]
                                    cats = [self.engine.convert_text(c) for c in cats]
                                    if cats and not cat_list:
                                        cat_list = cats

                                    vals = [v.text.strip() for v in ser.xpath("./*[local-name()='val']//*[local-name()='v']") if v.text and v.text.strip()]
                                    series_list.append((ser_name, vals))

                                chart_content_parts = []
                                if title_str:
                                    chart_content_parts.append(f'<span style="font-weight:bold; color:#92400e;">图表标题：{_html.escape(title_str)}</span>')

                                if series_list:
                                    tbl_parts = ['<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse; border-color:#fcd34d; margin-top:6px; background-color:#ffffff; font-size:10pt; width:96%;">']
                                    tbl_parts.append('<tr style="background-color:#fef3c7;">')
                                    tbl_parts.append('<th style="border:1px solid #fcd34d; padding:4px 8px; text-align:left; color:#92400e;">分类 / 维度</th>')
                                    for ser_name, _ in series_list:
                                        tbl_parts.append(f'<th style="border:1px solid #fcd34d; padding:4px 8px; text-align:right; color:#92400e;">{_html.escape(ser_name)}</th>')
                                    tbl_parts.append('</tr>')

                                    if cat_list:
                                        for idx, cat_name in enumerate(cat_list):
                                            tbl_parts.append('<tr>')
                                            tbl_parts.append(f'<td style="border:1px solid #fcd34d; padding:4px 8px; background-color:#fffbeb;">{_html.escape(cat_name)}</td>')
                                            for ser_name, vals in series_list:
                                                v_str = vals[idx] if idx < len(vals) else "-"
                                                tbl_parts.append(f'<td style="border:1px solid #fcd34d; padding:4px 8px; text-align:right;">{_html.escape(v_str)}</td>')
                                            tbl_parts.append('</tr>')
                                    else:
                                        for ser_name, vals in series_list:
                                            tbl_parts.append('<tr>')
                                            tbl_parts.append(f'<td style="border:1px solid #fcd34d; padding:4px 8px; background-color:#fffbeb;">{_html.escape(ser_name)}</td>')
                                            v_text = ", ".join(vals) if vals else "-"
                                            tbl_parts.append(f'<td style="border:1px solid #fcd34d; padding:4px 8px;" colspan="{len(series_list)}">{_html.escape(v_text)}</td>')
                                            tbl_parts.append('</tr>')
                                    tbl_parts.append('</table>')
                                    chart_content_parts.append("".join(tbl_parts))
                                else:
                                    raw_texts = [t.text.strip() for t in c_root.xpath('.//*[local-name()="t"] | .//*[local-name()="v"]') if t.text and t.text.strip()]
                                    if raw_texts:
                                        joined_raw = " | ".join(_html.escape(self.engine.convert_text(t)) for t in raw_texts[:10])
                                        chart_content_parts.append(f'<div style="margin-top:4px;">{joined_raw}</div>')

                                inner_chart = "<br/>".join(chart_content_parts) if chart_content_parts else "图表数据"
                                parts.append(
                                    '<div style="background-color:#fffbeb; border:1px solid #fcd34d; padding:8px 12px; margin:10px 0;">'
                                    '<b style="color:#b45309;">【图表 / Chart 数据预览】</b><br/>'
                                    + inner_chart + "</div>"
                                )

            elif ext == ".pdf":
                doc = pymupdf.open(file_path)
                for p_idx, page in enumerate(doc):
                    parts.append(
                        f'<div style="background-color:#f1f5f9; padding:4px 8px; margin:10px 0; font-weight:bold; color:#334155;">📄 第 {p_idx + 1} 页</div>'
                    )
                    p_dict = page.get_text("dict") if hasattr(page, "get_text") else None
                    if not isinstance(p_dict, dict) or not p_dict.get("blocks"):
                        raw_text = page.get_text() if hasattr(page, "get_text") else ""
                        if isinstance(raw_text, list):
                            lines = [item[4] for item in raw_text if len(item) > 4 and isinstance(item[4], str)]
                            raw_text = "\n".join(lines)
                        if isinstance(raw_text, str) and raw_text.strip():
                            for line in raw_text.splitlines():
                                l_s = line.strip()
                                if l_s:
                                    parts.append(f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(self.engine.convert_text(l_s))}</p>')
                        continue

                    # 分类 blocks: SmartArt/大字号/白色文字 vs 普通文本/表格行
                    normal_lines = []
                    for b in p_dict.get("blocks", []):
                        if b.get("type") != 0:
                            continue

                        # 自动识别大字号彩色背景块 (如 size >= 18 或 color == 0xffffff / 16777215 的 SmartArt 块)
                        is_smartart = False
                        for l in b.get("lines", []):
                            for s in l.get("spans", []):
                                sz = float(s.get("size", 12.0))
                                col = int(s.get("color", 0))
                                if sz >= 18.0 or col in (0xffffff, 16777215):
                                    is_smartart = True
                                    break
                            if is_smartart:
                                break

                        if is_smartart:
                            b_text_parts = []
                            for l in b.get("lines", []):
                                l_txt = "".join(ch for s in l.get("spans", []) for ch in s.get("text", "") if not ('\uf000' <= ch <= '\uf8ff')).strip()
                                if l_txt:
                                    b_text_parts.append(l_txt)
                            if b_text_parts:
                                b_text = "<br/>".join(_html.escape(self.engine.convert_text(t)) for t in b_text_parts)
                                parts.append(
                                    f'<div style="background-color:#eff6ff; border:1px solid #93c5fd; padding:8px;"><b style="color:#1d4ed8;">【SmartArt / 图形文本】</b><br/>{b_text}</div>'
                                )
                        else:
                            for l in b.get("lines", []):
                                l_txt = "".join(ch for s in l.get("spans", []) for ch in s.get("text", "") if not ('\uf000' <= ch <= '\uf8ff')).strip()
                                if l_txt:
                                    normal_lines.append({
                                        "text": l_txt,
                                        "bbox": l.get("bbox", (0, 0, 0, 0)),
                                        "spans": l.get("spans", [])
                                    })

                    if not normal_lines:
                        # 扫描版 PDF 页面：通过 OCR 渲染富文本内容
                        try:
                            import tempfile
                            pix = page.get_pixmap(dpi=150)
                            tf = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                            tf_path = tf.name
                            tf.close()
                            pix.save(tf_path)
                            ocr_res = self._ocr_via_windows_media(tf_path)
                            try:
                                os.remove(tf_path)
                            except Exception:
                                pass
                            if ocr_res and ocr_res.strip():
                                for line in ocr_res.splitlines():
                                    l_s = line.strip()
                                    if l_s:
                                        parts.append(f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(self.engine.convert_text(l_s))}</p>')
                        except Exception:
                            pass
                        continue

                    # 自动识别页面中的多列网格/表格行（同一 y 坐标附近的多个水平分布块）
                    normal_lines.sort(key=lambda item: (item["bbox"][1], item["bbox"][0]))
                    clusters = []
                    curr_cluster = []
                    for item in normal_lines:
                        if not curr_cluster:
                            curr_cluster.append(item)
                        elif abs(item["bbox"][1] - curr_cluster[0]["bbox"][1]) <= 6.0:
                            curr_cluster.append(item)
                        else:
                            clusters.append(curr_cluster)
                            curr_cluster = [item]
                    if curr_cluster:
                        clusters.append(curr_cluster)

                    table_accumulator = []

                    def flush_pdf_table():
                        if not table_accumulator:
                            return
                        tbl_html = ['<table border="1" cellspacing="0" cellpadding="5" style="border-collapse:collapse; border-color:#cbd5e1; margin:10px 0; width:98%;">']
                        for r_idx, row_cells in enumerate(table_accumulator):
                            tbl_html.append('<tr>')
                            for cell in row_cells:
                                bg = 'background-color:#f8fafc; font-weight:bold;' if r_idx == 0 else 'background-color:#ffffff;'
                                tbl_html.append(f'<td style="{bg} padding:5px; border:1px solid #cbd5e1;">{_html.escape(self.engine.convert_text(cell))}</td>')
                            tbl_html.append('</tr>')
                        tbl_html.append('</table>')
                        parts.append("".join(tbl_html))
                        table_accumulator.clear()

                    for cluster in clusters:
                        cluster.sort(key=lambda x: x["bbox"][0])
                        if len(cluster) >= 2:
                            cells = [x["text"] for x in cluster]
                            table_accumulator.append(cells)
                        else:
                            flush_pdf_table()
                            p_text = cluster[0]["text"]
                            parts.append(
                                f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(self.engine.convert_text(p_text))}</p>'
                            )
                    flush_pdf_table()
                doc.close()

            elif ext == ".odt":
                odt_handled = False
                if zipfile.is_zipfile(file_path):
                    try:
                        with zipfile.ZipFile(file_path, "r") as z:
                            if "content.xml" in z.namelist():
                                from lxml import etree
                                root = etree.fromstring(z.read("content.xml"))
                                for elem in root.iter():
                                    lname = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                                    if lname in ("p", "h"):
                                        p_txt = "".join(elem.itertext()).strip()
                                        p_txt = "".join(ch for ch in p_txt if not ('\uf000' <= ch <= '\uf8ff')).strip()
                                        if p_txt:
                                            parts.append(f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(self.engine.convert_text(p_txt))}</p>')
                                            odt_handled = True
                                    elif lname == "table":
                                        rows_html = []
                                        rows = [r for r in elem.iter() if (r.tag.split("}")[-1] if "}" in r.tag else r.tag) in ("table-row", "row")]
                                        for r_idx, r in enumerate(rows):
                                            cells_html = []
                                            bg = 'background-color:#f8fafc; font-weight:bold;' if r_idx == 0 else 'background-color:#ffffff;'
                                            cells = [c for c in r.iter() if (c.tag.split("}")[-1] if "}" in c.tag else c.tag) in ("table-cell", "cell")]
                                            for cell in cells:
                                                c_txt = "".join(cell.itertext()).strip()
                                                c_txt = "".join(ch for ch in c_txt if not ('\uf000' <= ch <= '\uf8ff')).strip()
                                                cells_html.append(f'<td style="{bg} padding:6px; border:1px solid #cbd5e1;">{_html.escape(self.engine.convert_text(c_txt))}</td>')
                                            if cells_html:
                                                rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
                                        if rows_html:
                                            parts.append(
                                                '<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; border-color:#cbd5e1; margin:10px 0; width:98%;">'
                                                + "".join(rows_html) + "</table>"
                                            )
                                            odt_handled = True
                    except Exception:
                        odt_handled = False
                if not odt_handled:
                    raw_txt = self._extract_odt_text(file_path) or self.extract_text(file_path)
                    if raw_txt:
                        for line in raw_txt.splitlines():
                            l_s = line.strip()
                            if l_s:
                                parts.append(f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(self.engine.convert_text(l_s))}</p>')

            elif ext == ".doc":
                raw_txt = self._extract_doc_text(file_path) or self.extract_text(file_path)
                if raw_txt:
                    for line in raw_txt.splitlines():
                        l_s = line.strip()
                        if l_s:
                            parts.append(f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(self.engine.convert_text(l_s))}</p>')

            elif ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"]:
                cached_txt = ""
                if os.path.exists(file_path + ".ocr.txt"):
                    try:
                        with open(file_path + ".ocr.txt", "r", encoding="utf-8") as f:
                            cached_txt = f.read().strip()
                    except Exception:
                        cached_txt = ""
                if not cached_txt:
                    cached_txt = self.extract_text(file_path)
                    if cached_txt:
                        cached_txt = self.engine.convert_text(cached_txt)
                if cached_txt:
                    for line in cached_txt.splitlines():
                        l_s = line.strip()
                        if l_s:
                            parts.append(f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(line)}</p>')
            else:
                raw_txt = self.extract_text(file_path)
                if raw_txt:
                    for line in raw_txt.splitlines():
                        l_s = line.strip()
                        if l_s:
                            parts.append(f'<p style="margin:6px 0; line-height:1.5; font-size:11pt;">{_html.escape(self.engine.convert_text(l_s))}</p>')
        except Exception:
            pass

        parts.append("</div>")
        return "".join(parts) if len(parts) > 2 else ""

    def render_page_images(self, file_path: str, dpi: int = 130, max_pages: int = 15) -> list[str]:
        """按需将转换后的文档（PDF/DOCX/DOC/ODT/图片）渲染为高清页面 PNG 路径列表"""
        if not file_path or not os.path.exists(file_path):
            return []
        ext = os.path.splitext(file_path)[1].lower()
        if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"]:
            return [os.path.abspath(file_path)]

        import tempfile
        out_dir = tempfile.mkdtemp(prefix="ChsToCht_Pages_")
        pdf_target = file_path if ext == ".pdf" else os.path.join(out_dir, "rendered_preview.pdf")

        if ext in [".docx", ".doc", ".odt"]:
            converted_to_pdf = False
            # 优先使用本机 Word COM 导出高保真 PDF (wdFormatPDF = 17)
            if ext in [".docx", ".doc", ".odt"]:
                try:
                    import win32com.client
                    import pythoncom
                    pythoncom.CoInitialize()
                    word = None
                    doc = None
                    try:
                        word = win32com.client.Dispatch("Word.Application")
                        word.Visible = False
                        word.DisplayAlerts = 0
                        doc = word.Documents.Open(
                            os.path.abspath(file_path),
                            ConfirmConversions=False,
                            ReadOnly=True
                        )
                        doc.SaveAs2(os.path.abspath(pdf_target), FileFormat=17)
                        doc.Close(False)
                        doc = None
                        converted_to_pdf = os.path.exists(pdf_target) and os.path.getsize(pdf_target) > 0
                    finally:
                        if doc is not None:
                            try:
                                doc.Close(False)
                            except Exception:
                                pass
                        if word is not None:
                            try:
                                word.Quit()
                            except Exception:
                                pass
                        try:
                            pythoncom.CoUninitialize()
                        except Exception:
                            pass
                except Exception:
                    converted_to_pdf = False

            # 回退使用 LibreOffice 转换为临时 PDF
            if not converted_to_pdf:
                try:
                    import subprocess
                    cmd = [
                        "soffice", "--headless", "--convert-to", "pdf",
                        os.path.abspath(file_path), "--outdir", out_dir
                    ]
                    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=40)
                    base_name = os.path.splitext(os.path.basename(file_path))[0]
                    lo_pdf = os.path.join(out_dir, f"{base_name}.pdf")
                    if os.path.exists(lo_pdf) and os.path.getsize(lo_pdf) > 0:
                        pdf_target = lo_pdf
                        converted_to_pdf = True
                except Exception:
                    converted_to_pdf = False

            if not converted_to_pdf:
                return []

        page_paths = []
        try:
            doc = pymupdf.open(pdf_target)
            count = min(len(doc), max_pages)
            for idx in range(count):
                pix = doc[idx].get_pixmap(dpi=dpi)
                img_path = os.path.join(out_dir, f"page_{idx + 1}.png")
                pix.save(img_path)
                page_paths.append(img_path)
            doc.close()
        except Exception:
            pass
        return page_paths

    def _generate_docx_preview(self, docx_path: str) -> str:
        """提取 docx 文本内容片段供前端极速预览，合并正文段落、表格内容、文字方块、SmartArt、图表文字"""
        try:
            preview_lines = []
            has_more = False

            # 1. 正文段落与表格内容
            try:
                doc = docx.Document(docx_path)
                for p in doc.paragraphs[:20]:  # 获取前 20 段提供给 UI 预览
                    text = p.text.strip()
                    if text:
                        preview_lines.append(text)
                if len(doc.paragraphs) > 20:
                    has_more = True

                # 表格内容（无论正文是否有段落都合并包含）
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                text = p.text.strip()
                                if text:
                                    preview_lines.append(text)
                                if len(preview_lines) >= 30:
                                    has_more = True
                                    break
                            if len(preview_lines) >= 30:
                                break
                        if len(preview_lines) >= 30:
                            break
                    if len(preview_lines) >= 30:
                        break
            except Exception:
                pass

            # 2. 文字方块、SmartArt、图表、页眉页脚
            extras = self._extract_docx_extra_texts(docx_path)
            for item in extras:
                if len(preview_lines) >= 40:
                    has_more = True
                    break
                preview_lines.append(item)

            if has_more or len(extras) > 40:
                preview_lines.append("\n... [更多内容详见导出的完整文件]")
            return "\n".join(preview_lines)
        except Exception:
            return ""

    def _get_output_path(self, file_path: str) -> str:
        d, f = os.path.split(file_path)
        name, ext = os.path.splitext(f)
        return os.path.join(d, f"{name}_繁体{ext}")

    def _process_docx(self, input_path: str, output_path: str, progress_cb):
        """原地保留样式的 DOCX 段落/表格文本节点替换，并无损转换文字方块、SmartArt、图表等 OpenXML 组件"""
        doc = docx.Document(input_path)
        total_steps = len(doc.paragraphs) + sum(len(t.rows) for t in doc.tables) + 1
        current = 0

        # 1. 遍历正文段落
        for p in doc.paragraphs:
            convert_paragraph_runs_safely(p, self.engine.convert_text)
            current += 1
            if progress_cb:
                progress_cb(int((current / total_steps) * 100))

        # 2. 遍历表格单元格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        convert_paragraph_runs_safely(p, self.engine.convert_text)
                current += 1
                if progress_cb:
                    progress_cb(int((current / total_steps) * 100))

        doc.save(output_path)

        # 3. 原地转换文字方块、SmartArt、图表、页眉页脚等 OpenXML 部件，同时补回未被序列化的附件
        if os.path.exists(output_path) and zipfile.is_zipfile(output_path):
            self._process_docx_xml_extras(output_path, input_path=input_path)

        if progress_cb:
            progress_cb(100)

    def _process_pdf(self, input_path: str, output_path: str, progress_cb):
        """PyMuPDF 高保真 Span 级原地替换：保留原字号、原颜色、原基线位置、表格线与 SmartArt 矢量背景框"""
        doc = pymupdf.open(input_path)
        total_pages = len(doc)

        # 查找本机可用的 Windows CJK 字体文件（优先等线 Deng.ttf，其次微软雅黑 msyh.ttc / 宋体 simsun.ttc）
        cjk_font_path = None
        for candidate in [
            r"C:\Windows\Fonts\Deng.ttf",
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\simsun.ttc",
        ]:
            if os.path.exists(candidate):
                cjk_font_path = candidate
                break

        for page_idx, page in enumerate(doc):
            page_dict = None
            try:
                res = page.get_text("dict")
                if isinstance(res, dict) and "blocks" in res:
                    page_dict = res
            except Exception:
                page_dict = None

            if page_dict is not None:
                # 真实 PDF 页面：使用 Span 级高保真原地无损替换
                if cjk_font_path:
                    try:
                        page.insert_font(fontname="F_CJK", fontfile=cjk_font_path)
                    except Exception:
                        pass

                text_blocks = [b for b in page_dict.get("blocks", []) if b.get("type") == 0]

                if not text_blocks:
                    # 纯扫描版 PDF 页面：通过高保真光栅化后执行 OCR 识别并覆盖重绘
                    try:
                        import tempfile
                        pix = page.get_pixmap(dpi=150)
                        tf_in = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                        tf_in_path = tf_in.name
                        tf_in.close()
                        tf_out = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                        tf_out_path = tf_out.name
                        tf_out.close()
                        pix.save(tf_in_path)
                        self._process_jpg(tf_in_path, tf_out_path, None)
                        if os.path.exists(tf_out_path):
                            page.clean_contents()
                            page.insert_image(page.rect, filename=tf_out_path)
                        try:
                            os.remove(tf_in_path)
                            os.remove(tf_out_path)
                            if os.path.exists(tf_out_path + ".ocr.txt"):
                                os.remove(tf_out_path + ".ocr.txt")
                        except Exception:
                            pass
                    except Exception:
                        pass

                # 收集跨 block 的折行组（解决窄表格单元格与 SmartArt 折行分词问题）
                groups = []
                for b in text_blocks:
                    spans = []
                    for l in b.get("lines", []):
                        for s in l.get("spans", []):
                            txt = s.get("text", "")
                            # 跳过纯空白、纯 ASCII 或 Word 列表符号私用区字符 (\uf0b7)
                            if txt and any('\u4e00' <= c <= '\u9fff' for c in txt):
                                spans.append(s)
                    if not spans:
                        continue

                    b_has_bullet = any(
                        any('\uf000' <= ch <= '\uf8ff' for ch in s.get("text", "")) or "•" in s.get("text", "") or "·" in s.get("text", "")
                        for l in b.get("lines", []) for s in l.get("spans", [])
                    )
                    first_txt = spans[0].get("text", "").strip()
                    last_txt = spans[-1].get("text", "").strip()

                    merged = False
                    if groups and not b_has_bullet:
                        prev_b, prev_spans = groups[-1][-1]
                        prev_last_txt = prev_spans[-1].get("text", "").strip()
                        x_diff = abs(b["bbox"][0] - prev_b["bbox"][0])
                        y_diff = b["bbox"][1] - prev_b["bbox"][3]
                        curr_cjk_len = sum(len("".join(c for c in s.get("text", "") if '\u4e00' <= c <= '\u9fff')) for s in spans)

                        if (
                            '\u4e00' <= prev_last_txt[-1] <= '\u9fff'
                            and '\u4e00' <= first_txt[0] <= '\u9fff'
                            and x_diff < 20.0
                            and 0 <= y_diff < 40.0
                            and curr_cjk_len <= 6
                        ):
                            groups[-1].append((b, spans))
                            merged = True
                    if not merged:
                        groups.append([(b, spans)])

                replacements = []
                for g in groups:
                    group_spans = [s for _, spans in g for s in spans]
                    orig_concat = "".join(s.get("text", "") for s in group_spans)
                    conv_concat = self.engine.convert_text(orig_concat)

                    # 按各 span 原始字符长度切片回填
                    if len(conv_concat) == len(orig_concat):
                        offset = 0
                        span_convs = []
                        for s in group_spans:
                            slen = len(s.get("text", ""))
                            span_convs.append(conv_concat[offset:offset + slen])
                            offset += slen
                    else:
                        span_convs = [self.engine.convert_text(s.get("text", "")) for s in group_spans]

                    for s, conv_txt in zip(group_spans, span_convs):
                        orig_s_txt = s.get("text", "")
                        # 转换前后完全一致则无需重绘，保留原字符矢量
                        if conv_txt == orig_s_txt:
                            continue

                        bbox = pymupdf.Rect(s["bbox"])
                        origin_pt = s.get("origin")
                        if origin_pt and len(origin_pt) == 2:
                            origin = pymupdf.Point(origin_pt[0], origin_pt[1])
                        else:
                            origin = pymupdf.Point(bbox.x0, bbox.y1 - 2)
                        fontsize = float(s.get("size", 12.0))
                        color_int = int(s.get("color", 0))
                        r_col = ((color_int >> 16) & 255) / 255.0
                        g_col = ((color_int >> 8) & 255) / 255.0
                        b_col = (color_int & 255) / 255.0
                        text_color = (r_col, g_col, b_col)

                        # 透明擦除旧文字，不填充白色背景
                        page.add_redact_annot(bbox, fill=False)
                        replacements.append((origin, conv_txt, fontsize, text_color))

                if replacements:
                    # 仅移除被标记的旧文字字符，100% 保留底层的图片、表格边框线和 SmartArt 彩色矢量框
                    try:
                        page.apply_redactions(
                            images=getattr(pymupdf, "PDF_REDACT_IMAGE_NONE", 0),
                            graphics=getattr(pymupdf, "PDF_REDACT_LINE_ART_NONE", 0)
                        )
                    except TypeError:
                        page.apply_redactions()

                    for origin, conv_txt, fontsize, text_color in replacements:
                        if cjk_font_path:
                            page.insert_text(
                                origin,
                                conv_txt,
                                fontsize=fontsize,
                                fontname="F_CJK",
                                fontfile=cjk_font_path,
                                color=text_color
                            )
                        else:
                            page.insert_text(
                                origin,
                                conv_txt,
                                fontsize=fontsize,
                                fontname="china-t",
                                color=text_color
                            )

                # 检查并处理页面内嵌入的大尺寸光栅图片（扫描图、图表等）
                try:
                    for img_info in page.get_images(full=True):
                        xref = img_info[0]
                        width = img_info[2]
                        height = img_info[3]
                        if width >= 150 and height >= 80:
                            img_dict = doc.extract_image(xref)
                            if img_dict and "image" in img_dict:
                                import tempfile
                                ext_name = img_dict.get("ext", "png")
                                itf_in = tempfile.NamedTemporaryFile(suffix=f".{ext_name}", delete=False)
                                itf_in.write(img_dict["image"])
                                itf_in_path = itf_in.name
                                itf_in.close()
                                itf_out = tempfile.NamedTemporaryFile(suffix=f".{ext_name}", delete=False)
                                itf_out_path = itf_out.name
                                itf_out.close()
                                self._process_jpg(itf_in_path, itf_out_path, None)
                                sidecar_f = itf_out_path + ".ocr.txt"
                                if os.path.exists(sidecar_f):
                                    with open(sidecar_f, "r", encoding="utf-8") as sf:
                                        c_txt = sf.read().strip()
                                    if c_txt:
                                        page.replace_image(xref, filename=itf_out_path)
                                try:
                                    os.remove(itf_in_path)
                                    os.remove(itf_out_path)
                                    if os.path.exists(sidecar_f):
                                        os.remove(sidecar_f)
                                except Exception:
                                    pass
                except Exception:
                    pass
            else:
                # 兼容单元测试 Mock 分支
                text_instances = page.get_text("blocks")
                for inst in text_instances:
                    if inst[6] == 0:
                        bbox = pymupdf.Rect(inst[:4])
                        orig_text = inst[4]
                        conv_text = self.engine.convert_text(orig_text)
                        page.add_redact_annot(bbox)
                        page.apply_redactions()
                        page.insert_textbox(bbox, conv_text, fontsize=10, fontname="china-t")

            if progress_cb and total_pages > 0:
                progress_cb(int(((page_idx + 1) / total_pages) * 100))

        doc.save(output_path)
        doc.close()

    def _process_via_libreoffice(self, input_path: str, output_path: str, progress_cb):
        """
        处理 .doc / .odt：
        调用本地无头 libreoffice 命令行将 .doc 转为临时的 .docx，
        经由 _process_docx 处理后，再用 libreoffice 转回目标格式。
        如果本机没有安装 soffice 则根据扩展名回退至原生 ODT 或 Word COM 处理。
        """
        import subprocess
        import tempfile
        import shutil

        try:
            if progress_cb:
                progress_cb(20)

            with tempfile.TemporaryDirectory() as tmp_dir:
                # 1. 转为临时 docx
                cmd_to_docx = [
                    "soffice", "--headless", "--convert-to", "docx",
                    input_path, "--outdir", tmp_dir
                ]
                subprocess.run(cmd_to_docx, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

                base_name = os.path.splitext(os.path.basename(input_path))[0]
                tmp_docx = os.path.join(tmp_dir, f"{base_name}.docx")
                tmp_docx_out = os.path.join(tmp_dir, f"{base_name}_converted.docx")

                if progress_cb:
                    progress_cb(50)

                # 2. 对中间 docx 进行替换
                self._process_docx(tmp_docx, tmp_docx_out, None)

                if progress_cb:
                    progress_cb(80)

                target_ext = os.path.splitext(input_path)[1][1:].lower()
                if target_ext == "doc":
                    # 重新转回 doc
                    cmd_back = [
                        "soffice", "--headless", "--convert-to", "doc",
                        tmp_docx_out, "--outdir", tmp_dir
                    ]
                    subprocess.run(cmd_back, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    final_tmp = os.path.join(tmp_dir, f"{base_name}_converted.doc")
                    shutil.copy2(final_tmp, output_path)
                elif target_ext == "odt":
                    cmd_back = [
                        "soffice", "--headless", "--convert-to", "odt",
                        tmp_docx_out, "--outdir", tmp_dir
                    ]
                    subprocess.run(cmd_back, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    final_tmp = os.path.join(tmp_dir, f"{base_name}_converted.odt")
                    shutil.copy2(final_tmp, output_path)

            if progress_cb:
                progress_cb(100)
        except (FileNotFoundError, OSError, subprocess.SubprocessError):
            ext = os.path.splitext(input_path)[1].lower()
            if ext == ".odt":
                self._process_odt_native(input_path, output_path, progress_cb)
            elif ext == ".doc":
                self._process_doc_via_word_com(input_path, output_path, progress_cb)
            else:
                raise

    def _ocr_boxes_via_windows_media(self, file_path: str) -> dict:
        """在未安装 Tesseract 环境下调用 Windows 原生 OcrEngine 获取行级文字与坐标框（支持横版图超分锐化、2D 网格智能空白行列分片与旋转扫描件自动方向识别）"""
        import subprocess
        import json
        import tempfile
        from PIL import ImageEnhance, ImageFilter, ImageOps

        if not os.path.exists(file_path):
            return {'lines': [], 'level': [], 'text': [], 'left': [], 'top': [], 'width': [], 'height': []}

        try:
            orig_img = Image.open(file_path).convert('RGB')
            orig_w, orig_h = orig_img.size
        except Exception:
            return {'lines': [], 'level': [], 'text': [], 'left': [], 'top': [], 'width': [], 'height': []}

        def _find_best_cut_y(gray_img, target_y, search_radius=80):
            gw, gh = gray_img.size
            y1 = max(1, target_y - search_radius)
            y2 = min(gh - 1, target_y + search_radius)
            if y1 >= y2:
                return target_y
            strip = gray_img.crop((0, y1, gw, y2))
            # 缩小宽度以极速计算每行像素方差
            sw = min(gw, 300)
            strip_small = strip.resize((sw, y2 - y1), Image.Resampling.BILINEAR)
            px = list(strip_small.get_flattened_data() if hasattr(strip_small, "get_flattened_data") else strip_small.getdata())
            best_y = target_y
            best_var = float('inf')
            for r_idx in range(y2 - y1):
                row = px[r_idx * sw:(r_idx + 1) * sw]
                if not row:
                    continue
                mean = sum(row) / sw
                var = sum((v - mean) ** 2 for v in row) / sw
                if var < best_var:
                    best_var = var
                    best_y = y1 + r_idx
            return best_y

        def _find_best_cut_x(gray_img, target_x, search_radius=80):
            gw, gh = gray_img.size
            x1 = max(1, target_x - search_radius)
            x2 = min(gw - 1, target_x + search_radius)
            if x1 >= x2:
                return target_x
            strip = gray_img.crop((x1, 0, x2, gh))
            sh = min(gh, 300)
            strip_small = strip.resize((x2 - x1, sh), Image.Resampling.BILINEAR)
            px = list(strip_small.get_flattened_data() if hasattr(strip_small, "get_flattened_data") else strip_small.getdata())
            cw = x2 - x1
            best_x = target_x
            best_var = float('inf')
            for c_idx in range(cw):
                col = [px[r_idx * cw + c_idx] for r_idx in range(sh)]
                if not col:
                    continue
                mean = sum(col) / sh
                var = sum((v - mean) ** 2 for v in col) / sh
                if var < best_var:
                    best_var = var
                    best_x = x1 + c_idx
            return best_x

        def _remove_table_grid_lines(img_obj):
            """移除长表格横竖细边框线，保留与之相交的文字笔画，防止紧贴格线的文字（如‘热干面’、‘老北京炸酱面’）被格线粘连忽略"""
            ww, wh = img_obj.size
            if ww < 100 or wh < 100:
                return img_obj
            out = img_obj.copy()
            px = out.load()
            h_marks = []
            for y in range(2, wh - 2):
                rs = None
                for x in range(ww):
                    r, g, b = px[x, y]
                    if r < 160 and g < 160 and b < 160:
                        if rs is None:
                            rs = x
                    else:
                        if rs is not None and (x - rs) > 140:
                            for xx in range(rs, x):
                                if px[xx, y - 2][0] >= 160 or px[xx, y + 2][0] >= 160:
                                    h_marks.append((xx, y))
                        rs = None
                if rs is not None and (ww - rs) > 140:
                    for xx in range(rs, ww):
                        if px[xx, y - 2][0] >= 160 or px[xx, y + 2][0] >= 160:
                            h_marks.append((xx, y))

            v_marks = []
            for x in range(2, ww - 2):
                rs = None
                for y in range(wh):
                    r, g, b = px[x, y]
                    if r < 160 and g < 160 and b < 160:
                        if rs is None:
                            rs = y
                    else:
                        if rs is not None and (y - rs) > 100:
                            for yy in range(rs, y):
                                if px[x - 2, yy][0] >= 160 or px[x + 2, yy][0] >= 160:
                                    v_marks.append((x, yy))
                        rs = None
                if rs is not None and (wh - rs) > 100:
                    for yy in range(rs, wh):
                        if px[x - 2, yy][0] >= 160 or px[x + 2, yy][0] >= 160:
                            v_marks.append((x, yy))

            for xx, yy in h_marks:
                px[xx, yy] = (255, 255, 255)
            for xx, yy in v_marks:
                px[xx, yy] = (255, 255, 255)
            return out

        def _build_tiles_for_image(base_img):
            bw, bh = base_img.size
            scale_factor = 1.0
            work_img = base_img
            # 自适应缩放到 Windows OcrEngine 最佳单页全宽区间 [1400, 2100]：
            # 1) 小图 (< 1050px) 超分放大，避免小字漏识；
            # 2) 超大图 (> 2100px，如 300DPI 高清 JPEG) 适度缩放到 2100px 宽，既避免字高超过 100px 拆成偏旁，又避免水平一刀切碎文字！
            if bw < 1050:
                scale_factor = min(2.2, 1600.0 / max(1, bw))
            elif bw > 2100 and bw <= bh * 1.4:
                scale_factor = 2100.0 / bw
            elif bw > bh and bh < 1050:
                scale_factor = min(2.0, 1400.0 / max(1, bh))
                if bw * scale_factor > 3600:
                    scale_factor = max(1.0, 3600.0 / bw)

            if abs(scale_factor - 1.0) > 0.03:
                work_img = base_img.resize(
                    (int(round(bw * scale_factor)), int(round(bh * scale_factor))),
                    Image.Resampling.LANCZOS
                )
            work_img = _remove_table_grid_lines(work_img)

            W, H = work_img.size
            gray = ImageOps.grayscale(work_img)
            tile_w = 1800
            tile_h = 1850
            overlap_x = 160
            overlap_y = 140

            y_ranges = []
            if H <= 2200:
                y_ranges.append((0, H))
            else:
                y = 0
                while y < H:
                    if y + tile_h >= H:
                        y_ranges.append((y, H))
                        break
                    cut_y = _find_best_cut_y(gray, y + tile_h, 90)
                    y_end = min(H, max(y + 400, cut_y + overlap_y // 2))
                    y_ranges.append((y, y_end))
                    y = max(y + 300, cut_y - overlap_y // 2)

            x_ranges = []
            if W <= 2250:
                x_ranges.append((0, W))
            else:
                x = 0
                while x < W:
                    if x + tile_w >= W:
                        x_ranges.append((x, W))
                        break
                    cut_x = _find_best_cut_x(gray, x + tile_w, 90)
                    x_end = min(W, max(x + 400, cut_x + overlap_x // 2))
                    x_ranges.append((x, x_end))
                    x = max(x + 300, cut_x - overlap_x // 2)

            tiles = []
            for y0, y1 in y_ranges:
                for x0, x1 in x_ranges:
                    tiles.append((x0, y0, work_img.crop((x0, y0, x1, y1))))
            return tiles, scale_factor

        def _run_ps_ocr_on_tiles(tiles_list):
            t_files = []
            try:
                for idx, (x0, y0, t_img) in enumerate(tiles_list):
                    tf = tempfile.NamedTemporaryFile(suffix=f"_tile_{idx}.png", delete=False)
                    tf_path = tf.name
                    tf.close()
                    t_img.save(tf_path)
                    t_files.append((tf_path, x0, y0, t_img.width, t_img.height))

                ps_paths = ", ".join(f"'{os.path.abspath(p).replace(chr(39), chr(39)*2)}'" for p, _, _, _, _ in t_files)
                ps_script = f"""
$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {{ $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.ContainsGenericParameters }})[0]
function Await($op, $t) {{
    $tTask = $asTaskGeneric.MakeGenericMethod($t).Invoke($null, @($op))
    if (-not $tTask.Wait(8000)) {{ return $null }}
    return $tTask.Result
}}
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$null = [Windows.Foundation.Rect, Windows.Foundation, ContentType = WindowsRuntime]

$lang = [Windows.Globalization.Language]::new('zh-Hans-CN')
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if (-not $engine) {{ $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }}
if (-not $engine) {{ exit 0 }}

$paths = @({ps_paths})
$allResults = @()
for ($i = 0; $i -lt $paths.Length; $i++) {{
    $p = $paths[$i]
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($p)) ([Windows.Storage.StorageFile])
    if (-not $file) {{ $allResults += @{{ band_idx = $i; text_angle = 0.0; lines = @() }}; continue }}
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    if (-not $stream) {{ $allResults += @{{ band_idx = $i; text_angle = 0.0; lines = @() }}; continue }}
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    if (-not $decoder) {{ $allResults += @{{ band_idx = $i; text_angle = 0.0; lines = @() }}; continue }}
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    if (-not $bitmap) {{ $allResults += @{{ band_idx = $i; text_angle = 0.0; lines = @() }}; continue }}
    $res = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    $angleVal = 0.0
    if ($res -and $null -ne $res.TextAngle) {{ $angleVal = [double]$res.TextAngle }}
    $linesData = @()
    if ($res -and $res.Lines) {{
        foreach ($line in $res.Lines) {{
            $words = @()
            if ($line.Words) {{
                foreach ($w in $line.Words) {{
                    $r = $w.BoundingRect
                    $words += @{{
                        text = $w.Text
                        left = [double]$r.X
                        top = [double]$r.Y
                        width = [double]$r.Width
                        height = [double]$r.Height
                    }}
                }}
            }}
            $lineText = ($line.Words | ForEach-Object {{ $_.Text }}) -join ''
            $linesData += @{{
                text = $lineText
                words = $words
            }}
        }}
    }}
    $allResults += @{{ band_idx = $i; text_angle = $angleVal; lines = $linesData }}
}}
$allResults | ConvertTo-Json -Depth 6 -Compress
"""
                res = subprocess.run(
                    ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                    capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=35
                )
                raw = res.stdout.strip()
                if not raw:
                    return []
                parsed = json.loads(raw)
                if isinstance(parsed, dict):
                    parsed = [parsed]
                return [
                    (t_files[i][1], t_files[i][2], t_files[i][3], t_files[i][4], parsed[i] if i < len(parsed) else {})
                    for i in range(len(t_files))
                ]
            finally:
                for p, _, _, _, _ in t_files:
                    try:
                        os.remove(p)
                    except Exception:
                        pass

        def _assemble_lines(tile_results, scale_factor):
            import math
            all_lines = []
            inv_scale = 1.0 / scale_factor
            for x_offset, y_offset, tile_w_px, tile_h_px, band_obj in tile_results:
                band_lines = band_obj.get('lines', []) if isinstance(band_obj, dict) else []
                text_angle = float(band_obj.get('text_angle', 0.0) or 0.0) if isinstance(band_obj, dict) else 0.0
                # Windows OcrEngine 当检测到非对齐行或斜排卡片产生非零 TextAngle 时，返回的 BoundingRect 是绕分片中心旋转后的坐标！
                # 必须先绕分片中心 (tile_w_px/2, tile_h_px/2) 逆旋转还原为真实图像像素坐标，否则纯文本行的“修改贴”擦除与重绘位置会出现几十像素的偏移！
                use_unrotate = abs(text_angle) > 0.01 and tile_w_px > 0 and tile_h_px > 0
                theta = math.radians(text_angle) if use_unrotate else 0.0
                cos_t = math.cos(theta) if use_unrotate else 1.0
                sin_t = math.sin(theta) if use_unrotate else 0.0
                cx_tile = tile_w_px / 2.0
                cy_tile = tile_h_px / 2.0

                for l in band_lines:
                    words = l.get('words', [])
                    scaled_words = []
                    for w in words:
                        wx = float(w.get('left', 0.0))
                        wy = float(w.get('top', 0.0))
                        ww = float(w.get('width', 10.0))
                        wh = float(w.get('height', 10.0))
                        if use_unrotate:
                            px = wx + ww / 2.0 - cx_tile
                            py = wy + wh / 2.0 - cy_tile
                            wx = cx_tile + px * cos_t - py * sin_t - ww / 2.0
                            wy = cy_tile + px * sin_t + py * cos_t - wh / 2.0
                        scaled_words.append({
                            'text': w.get('text', ''),
                            'left': int(round((wx + x_offset) * inv_scale)),
                            'top': int(round((wy + y_offset) * inv_scale)),
                            'width': max(1, int(round(ww * inv_scale))),
                            'height': max(1, int(round(wh * inv_scale))),
                        })
                    line_text = l.get('text', '')
                    if not scaled_words or not line_text.strip():
                        continue
                    line_top = scaled_words[0]['top']
                    line_left = scaled_words[0]['left']

                    # 跨分片重叠区域去重
                    is_dup = False
                    for prev in all_lines:
                        prev_words = prev.get('words', [])
                        if prev_words and prev.get('text') == line_text:
                            prev_top = prev_words[0]['top']
                            prev_left = prev_words[0]['left']
                            prev_h = max(25, prev_words[0].get('height', 25))
                            if abs(prev_top - line_top) < prev_h * 0.85 and abs(prev_left - line_left) < prev_h * 1.2:
                                is_dup = True
                                break
                    if not is_dup:
                        all_lines.append({
                            'text': line_text,
                            'words': scaled_words
                        })
            return all_lines

        try:
            tiles_0, scale_0 = _build_tiles_for_image(orig_img)
            res_0 = _run_ps_ocr_on_tiles(tiles_0)
            all_lines = _assemble_lines(res_0, scale_0)
            best_rotation = 0

            # 如果输入为横版图（宽 > 高），且 0° 方向几乎未识别出有效中文行（可能是旋转了 90°/270° 的竖版文档扫描件），自动探测旋转方向
            txt_0 = "".join(l.get('text', '') for l in all_lines)
            cjk_0 = sum(1 for ch in txt_0 if '\u4e00' <= ch <= '\u9fff')
            avg_len_0 = len(txt_0) / max(1, len(all_lines))
            if orig_w > orig_h and (cjk_0 < 25 or avg_len_0 < 2.2):
                best_cjk = cjk_0
                for rot_angle in (270, 90):
                    rot_img = orig_img.rotate(rot_angle, expand=True)
                    tiles_r, scale_r = _build_tiles_for_image(rot_img)
                    res_r = _run_ps_ocr_on_tiles(tiles_r)
                    lines_r = _assemble_lines(res_r, scale_r)
                    txt_r = "".join(l.get('text', '') for l in lines_r)
                    cjk_r = sum(1 for ch in txt_r if '\u4e00' <= ch <= '\u9fff')
                    avg_len_r = len(txt_r) / max(1, len(lines_r))
                    if cjk_r > best_cjk * 1.8 and avg_len_r >= 3.0:
                        best_cjk = cjk_r
                        all_lines = lines_r
                        best_rotation = rot_angle

            # 二次精细补扫（Pass 2）：针对 SmartArt 卡片、图文混排中因“首行项目符号缩进（右移非对齐）”而单字成行被漏扫的块（如：温/州猪/脏粉、腊/味双/拼饭、广/味腊/肠煲/仔饭、湖/南腊/肉炒/香干）
            try:
                active_img = orig_img.rotate(best_rotation, expand=True) if best_rotation in (90, 270) else orig_img
                aw, ah = active_img.size
                # 收集短中文行（1~3字）并按垂直对齐关系聚类为多行短文本块
                short_cjk = []
                for idx_l, l_obj in enumerate(all_lines):
                    t_str = l_obj.get('text', '').strip()
                    w_list = l_obj.get('words', [])
                    cjk_chars = [c for c in t_str if '\u4e00' <= c <= '\u9fff']
                    if 1 <= len(cjk_chars) <= 3 and w_list:
                        x0 = min(w['left'] for w in w_list)
                        y0 = min(w['top'] for w in w_list)
                        x1 = max(w['left'] + w['width'] for w in w_list)
                        y1 = max(w['top'] + w['height'] for w in w_list)
                        short_cjk.append({
                            'line_idx': idx_l,
                            'text': t_str,
                            'words': w_list,
                            'left': x0,
                            'top': y0,
                            'right': x1,
                            'bottom': y1,
                            'width': max(1, x1 - x0),
                            'height': max(1, y1 - y0),
                        })

                # 按列分桶（以字高的 2 倍为列宽阈值）并在每列内严格按 top 从小到大排序，保证同一张卡片的所有行排在一起
                col_Groups = []
                for item in sorted(short_cjk, key=lambda it: it['left']):
                    placed = False
                    for grp in col_Groups:
                        ref = grp[0]
                        if abs(item['left'] - ref['left']) < max(ref['height'], item['height']) * 1.8:
                            grp.append(item)
                            placed = True
                            break
                    if not placed:
                        col_Groups.append([item])
                short_cjk = []
                for grp in col_Groups:
                    grp.sort(key=lambda it: (it['top'], it['left']))
                    short_cjk.extend(grp)
                used_s = [False] * len(short_cjk)
                blocks = []
                for i in range(len(short_cjk)):
                    if used_s[i]:
                        continue
                    cur_b = [short_cjk[i]]
                    used_s[i] = True
                    for j in range(i + 1, len(short_cjk)):
                        if used_s[j]:
                            continue
                        cand = short_cjk[j]
                        last = cur_b[-1]
                        # 同一行的碎片（如 "脏'" 与 "包"，间距不超过 1.8 个字高）合并到同一逻辑行
                        if abs(cand['top'] - last['top']) < last['height'] * 0.45 and 0 <= (cand['left'] - last['right']) < last['height'] * 1.4:
                            last['words'] = last['words'] + cand['words']
                            last['text'] = last['text'] + cand['text']
                            last['right'] = max(last['right'], cand['right'])
                            last['bottom'] = max(last['bottom'], cand['bottom'])
                            last['width'] = last['right'] - last['left']
                            all_lines[last['line_idx']]['words'] = last['words']
                            all_lines[last['line_idx']]['text'] = last['text']
                            all_lines[cand['line_idx']]['words'] = []
                            all_lines[cand['line_idx']]['text'] = ''
                            used_s[j] = True
                            continue
                        # 垂直方向紧邻且左边界对齐（或属于同一卡片列：cand 在 last 正下方，或者 first_r 是单字右缩进且 cand 是其正下方左移 1 个字宽的主体行）
                        is_aligned_col = abs(cand['left'] - last['left']) < last['height'] * 0.65
                        is_indented_head = (
                            len(cur_b) == 1
                            and len([c for c in last['text'] if '\u4e00' <= c <= '\u9fff']) == 1
                            and (last['height'] * 0.5 <= (last['left'] - cand['left']) <= last['height'] * 1.8)
                        )
                        if (is_aligned_col or is_indented_head) and 0 < (cand['top'] - last['top']) < last['height'] * 2.1:
                            cur_b.append(cand)
                            used_s[j] = True
                    if len(cur_b) >= 2 or (len(cur_b) == 1 and len([c for c in cur_b[0]['text'] if '\u4e00' <= c <= '\u9fff']) == 2 and cur_b[0]['height'] >= 60):
                        blocks.append(cur_b)

                # 检查每个短块上方是否存在未对齐（右缩进）的首行孤立字，若有则与当前块各行横向拼接成单行条带送入 OCR
                strip_entries = []
                for b in blocks:
                    first_r = b[0]
                    # 若块的第一行已经是 1 个字（说明首行缩进字可能已被识别为单字行），则从第 2 行计算列左边界
                    if len(b) >= 2 and len([c for c in first_r['text'] if '\u4e00' <= c <= '\u9fff']) == 1 and first_r['left'] > b[1]['left'] + b[1]['height'] * 0.4:
                        col_left = b[1]['left']
                        col_right = max(r['right'] for r in b[1:])
                        col_h = b[1]['height']
                        head_box = (
                            max(0, first_r['left'] - 8),
                            max(0, first_r['top'] - 8),
                            min(aw, first_r['right'] + 12),
                            min(ah, first_r['bottom'] + 8),
                        )
                        body_rows = b[1:]
                        obsolete_line_indices = [r['line_idx'] for r in b]
                    else:
                        col_left = first_r['left']
                        col_right = max(r['right'] for r in b)
                        col_h = first_r['height']
                        char_w = max(col_h, int(round(first_r['width'] / max(1, len(first_r['text'])))))
                        # 搜索第一行正上方偏右（首行项目符号缩进位）的区域
                        hx1 = max(0, int(round(col_left + char_w * 0.85)))
                        hx2 = min(aw, int(round(col_right + char_w * 0.35)))
                        hy1 = max(0, int(round(first_r['top'] - col_h * 1.48)))
                        hy2 = max(0, int(round(first_r['top'] - col_h * 0.18)))
                        if hx2 - hx1 < 20 or hy2 - hy1 < 20:
                            continue
                        h_crop = active_img.crop((hx1, hy1, hx2, hy2))
                        h_px = list(h_crop.get_flattened_data() if hasattr(h_crop, "get_flattened_data") else h_crop.getdata())
                        bg_sample = h_crop.getpixel((0, 0))
                        fg_pts = [
                            (idx_p % h_crop.width, idx_p // h_crop.width)
                            for idx_p, p in enumerate(h_px)
                            if (p[0] - bg_sample[0]) ** 2 + (p[1] - bg_sample[1]) ** 2 + (p[2] - bg_sample[2]) ** 2 > 3600
                        ]
                        fg_ratio = len(fg_pts) / max(1, len(h_px))
                        if fg_ratio < 0.035 or len(fg_pts) < 40:
                            continue
                        min_fx = min(pt[0] for pt in fg_pts)
                        max_fx = max(pt[0] for pt in fg_pts)
                        min_fy = min(pt[1] for pt in fg_pts)
                        max_fy = max(pt[1] for pt in fg_pts)
                        head_box = (
                            max(0, hx1 + min_fx - 8),
                            max(0, hy1 + min_fy - 8),
                            min(aw, hx1 + max_fx + 12),
                            min(ah, hy1 + max_fy + 8),
                        )
                        body_rows = b
                        obsolete_line_indices = [r['line_idx'] for r in b]
                        # 同时检查 all_lines 中是否已有落在该 head_box 区域内的孤立单字行，将其一并加入 obsolete_line_indices 防止重复输出
                        for extra_idx, extra_l in enumerate(all_lines):
                            ew = extra_l.get('words', [])
                            if ew and extra_idx not in obsolete_line_indices:
                                ex0, ey0 = ew[0]['left'], ew[0]['top']
                                if (head_box[0] - 35) <= ex0 <= (head_box[2] + 35) and (head_box[1] - 35) <= ey0 <= (head_box[3] + 35):
                                    obsolete_line_indices.append(extra_idx)

                    # 组装 [head_box] + body_rows 的分片列表
                    segs = [head_box]
                    for r in body_rows:
                        segs.append((
                            max(0, r['left'] - 6),
                            max(0, r['top'] - 6),
                            min(aw, r['right'] + 10),
                            min(ah, r['bottom'] + 6),
                        ))
                    strip_entries.append((segs, obsolete_line_indices))

                if strip_entries:
                    # 将所有待补扫的块合成为一张虚拟水平长条图，每个分片之间保留 14px 原背景色间距防止笔画粘连
                    row_step = max(240, max((s[3] - s[1]) for segs, _ in strip_entries for s in segs) + 80)
                    canvas_w = max(1200, max(sum((s[2] - s[0]) + 14 for s in segs) + 140 for segs, _ in strip_entries))
                    canvas_h = max(400, len(strip_entries) * row_step + 60)
                    bg_fill = active_img.getpixel((strip_entries[0][0][0][0], strip_entries[0][0][0][1]))
                    strip_canvas = Image.new('RGB', (canvas_w, canvas_h), bg_fill)
                    entry_placements = []
                    for e_idx, (segs, obs_indices) in enumerate(strip_entries):
                        y_cursor = 30 + e_idx * row_step
                        x_cursor = 50
                        seg_map = []
                        for sx1, sy1, sx2, sy2 in segs:
                            scrop = active_img.crop((sx1, sy1, sx2, sy2))
                            strip_canvas.paste(scrop, (x_cursor, y_cursor))
                            seg_map.append((x_cursor, x_cursor + scrop.width, y_cursor, sx1, sy1, sx2, sy2))
                            x_cursor += scrop.width
                        entry_placements.append((y_cursor, y_cursor + (segs[0][3] - segs[0][1]) + 40, seg_map, obs_indices))

                    strip_res = _run_ps_ocr_on_tiles([(0, 0, strip_canvas)])
                    if strip_res:
                        _, _, s_tw, s_th, s_obj = strip_res[0]
                        s_lines = s_obj.get('lines', []) if isinstance(s_obj, dict) else []
                        s_angle = float(s_obj.get('text_angle', 0.0) or 0.0) if isinstance(s_obj, dict) else 0.0
                        s_unrot = abs(s_angle) > 0.01 and s_tw > 0 and s_th > 0
                        import math
                        s_theta = math.radians(s_angle) if s_unrot else 0.0
                        s_cos = math.cos(s_theta) if s_unrot else 1.0
                        s_sin = math.sin(s_theta) if s_unrot else 0.0
                        s_cx = s_tw / 2.0
                        s_cy = s_th / 2.0
                        for s_line in s_lines:
                            s_words = s_line.get('words', [])
                            s_text = s_line.get('text', '').strip()
                            if not s_words or not s_text:
                                continue
                            w0_x = float(s_words[0].get('left', 0.0))
                            w0_y = float(s_words[0].get('top', 0.0))
                            w0_w = float(s_words[0].get('width', 10.0))
                            w0_h = float(s_words[0].get('height', 10.0))
                            if s_unrot:
                                px0 = w0_x + w0_w / 2.0 - s_cx
                                py0 = w0_y + w0_h / 2.0 - s_cy
                                wy = s_cy + px0 * s_sin + py0 * s_cos - w0_h / 2.0
                            else:
                                wy = w0_y
                            for p_idx, (y_min, y_max, seg_map, obs_indices) in enumerate(entry_placements):
                                if (y_min - 30) <= wy <= (y_max + 40):
                                    mapped_words = []
                                    for sw in s_words:
                                        w_text = sw.get('text', '')
                                        # 去除被误识为顿号或斜杠的噪声符号
                                        if not any('\u4e00' <= c <= '\u9fff' or c.isalnum() for c in w_text):
                                            continue
                                        sw_x = float(sw.get('left', 0.0))
                                        sw_y = float(sw.get('top', 0.0))
                                        sw_w = float(sw.get('width', 10.0))
                                        sw_h = float(sw.get('height', 10.0))
                                        if s_unrot:
                                            px_s = sw_x + sw_w / 2.0 - s_cx
                                            py_s = sw_y + sw_h / 2.0 - s_cy
                                            sw_x = s_cx + px_s * s_cos - py_s * s_sin - sw_w / 2.0
                                            sw_y = s_cy + px_s * s_sin + py_s * s_cos - sw_h / 2.0
                                        wx_center = sw_x + sw_w / 2.0
                                        chosen_seg = seg_map[-1]
                                        for sm in seg_map:
                                            if sm[0] - 15 <= wx_center <= sm[1] + 15:
                                                chosen_seg = sm
                                                break
                                        cx0, cx1, cy0, orig_sx1, orig_sy1, orig_sx2, orig_sy2 = chosen_seg
                                        orig_left = max(orig_sx1, min(orig_sx2 - 10, int(round(orig_sx1 + (sw_x - cx0)))))
                                        orig_top = max(orig_sy1, min(orig_sy2 - 10, int(round(orig_sy1 + (sw_y - cy0)))))
                                        orig_w_box = min(orig_sx2 - orig_left, max(12, int(round(sw_w))))
                                        orig_h_box = min(orig_sy2 - orig_top, max(12, int(round(sw_h))))
                                        mapped_words.append({
                                            'text': w_text,
                                            'left': orig_left,
                                            'top': orig_top,
                                            'width': orig_w_box,
                                            'height': orig_h_box,
                                        })
                                    if mapped_words:
                                        for oi in obs_indices:
                                            if 0 <= oi < len(all_lines):
                                                all_lines[oi]['words'] = []
                                                all_lines[oi]['text'] = ''
                                        # 若同一个 strip_entry 被拆成多段返回，自动合并到同一行
                                        existing_target = None
                                        for al in all_lines:
                                            if al.get('_strip_entry_idx') == p_idx:
                                                existing_target = al
                                                break
                                        if existing_target is not None:
                                            existing_target['words'].extend(mapped_words)
                                            existing_target['text'] = "".join(mw['text'] for mw in existing_target['words'])
                                        else:
                                            all_lines.append({
                                                'text': "".join(mw['text'] for mw in mapped_words),
                                                'words': mapped_words,
                                                '_strip_entry_idx': p_idx,
                                            })
                                    break
                    all_lines = [l for l in all_lines if l.get('words') and l.get('text', '').strip()]
            except Exception:
                pass

            # 清理同一单元格/同一段落末尾单独折行的 1~2 字尾行（如“风干手撕牛”+“肉”、“老北京炸酱”+“面”、“终究舍不得离”+“开。”），使其合并为完整逻辑行以实现上下文精准消歧（如“炸酱面 -> 炸醬麵”）
            try:
                merged_flags = [False] * len(all_lines)
                for i_l, line_a in enumerate(all_lines):
                    if merged_flags[i_l]:
                        continue
                    wa = line_a.get('words', [])
                    ta = line_a.get('text', '').strip()
                    if not wa or not ta:
                        continue
                    ax0 = wa[0]['left']
                    ay0 = wa[0]['top']
                    ax1 = max(w['left'] + w['width'] for w in wa)
                    ay1 = max(w['top'] + w['height'] for w in wa)
                    aw_box = max(1, ax1 - ax0)
                    ah_box = max(1, ay1 - ay0)
                    for j_l, line_b in enumerate(all_lines):
                        if i_l == j_l or merged_flags[j_l]:
                            continue
                        wb = line_b.get('words', [])
                        tb = line_b.get('text', '').strip()
                        if not wb or not tb or len(tb) > 2:
                            continue
                        bx0 = wb[0]['left']
                        by0 = wb[0]['top']
                        # 如果短尾行在当前行正下方 1 行以内，且水平方向处于当前行左端附近（包括单元格左对齐比首行项目符号缩进更靠左的情况）
                        if 0 < (by0 - ay0) < ah_box * 3.8 and (ax0 - ah_box * 4.5) <= bx0 <= (ax0 + min(aw_box, ah_box * 4.5)):
                            line_a['words'] = wa + wb
                            line_a['text'] = ta + tb
                            merged_flags[j_l] = True
                            break
                all_lines = [l for idx, l in enumerate(all_lines) if not merged_flags[idx] and l.get('words')]
            except Exception:
                pass

            flattened_words = [w for item in all_lines for w in item.get('words', [])]
            return {
                'lines': all_lines,
                'rotation': best_rotation,
                'level': [1] * len(flattened_words),
                'text': [w.get('text', '') for w in flattened_words],
                'left': [int(w.get('left', 0)) for w in flattened_words],
                'top': [int(w.get('top', 0)) for w in flattened_words],
                'width': [int(w.get('width', 10)) for w in flattened_words],
                'height': [int(w.get('height', 10)) for w in flattened_words],
            }
        except Exception:
            return {'lines': [], 'rotation': 0, 'level': [], 'text': [], 'left': [], 'top': [], 'width': [], 'height': []}

    def _process_jpg(self, input_path: str, output_path: str, progress_cb):
        """利用 Tesseract OCR（或 Windows 原生 OCR 兜底）提取文字并用繁体字覆盖重绘（纯 Python 像素采样，零 numpy 依赖）"""
        image = Image.open(input_path).convert('RGB')

        def sample_colors(pil_img, x, y, w, h):
            if pil_img is None:
                return (255, 255, 255), (0, 0, 0)
            try:
                W, H = pil_img.size
                if not (isinstance(W, int) and isinstance(H, int)):
                    return (255, 255, 255), (0, 0, 0)
                x1 = max(0, x - 2)
                y1 = max(0, y - 2)
                x2 = min(W, x + w + 2)
                y2 = min(H, y + h + 2)
                if x2 <= x1 or y2 <= y1:
                    return (255, 255, 255), (0, 0, 0)

                perim = []
                for px_x in range(x1, x2):
                    perim.append(pil_img.getpixel((px_x, y1)))
                    perim.append(pil_img.getpixel((px_x, max(y1, y2 - 1))))
                for px_y in range(y1, y2):
                    perim.append(pil_img.getpixel((x1, px_y)))
                    perim.append(pil_img.getpixel((max(x1, x2 - 1), px_y)))

                if not perim or not isinstance(perim[0], (tuple, list)) or len(perim[0]) < 3:
                    return (255, 255, 255), (0, 0, 0)
                mid = len(perim) // 2
                bg = (
                    int(sorted(p[0] for p in perim)[mid]),
                    int(sorted(p[1] for p in perim)[mid]),
                    int(sorted(p[2] for p in perim)[mid]),
                )
                bg_lum = 0.299 * bg[0] + 0.587 * bg[1] + 0.114 * bg[2]

                ix1, iy1 = max(0, x), max(0, y)
                ix2, iy2 = min(W, x + w), min(H, y + h)
                if ix2 > ix1 and iy2 > iy1:
                    cropped = pil_img.crop((ix1, iy1, ix2, iy2))
                    inner_pixels = list(cropped.get_flattened_data() if hasattr(cropped, "get_flattened_data") else cropped.getdata())
                else:
                    inner_pixels = []

                if inner_pixels and isinstance(inner_pixels[0], (tuple, list)) and len(inner_pixels[0]) >= 3:
                    inner_pixels.sort(key=lambda p: (p[0] - bg[0]) ** 2 + (p[1] - bg[1]) ** 2 + (p[2] - bg[2]) ** 2, reverse=True)
                    top_k = max(1, int(len(inner_pixels) * 0.2))
                    farthest = inner_pixels[:top_k]
                    f_mid = len(farthest) // 2
                    fg = (
                        int(sorted(p[0] for p in farthest)[f_mid]),
                        int(sorted(p[1] for p in farthest)[f_mid]),
                        int(sorted(p[2] for p in farthest)[f_mid]),
                    )
                    fg_lum = 0.299 * fg[0] + 0.587 * fg[1] + 0.114 * fg[2]
                    if bg_lum > 140 and fg_lum > 130:
                        fg = (0, 0, 0)
                    elif bg_lum <= 140 and fg_lum <= 130:
                        fg = (255, 255, 255)
                else:
                    fg = (0, 0, 0) if bg_lum > 140 else (255, 255, 255)
                return bg, fg
            except Exception:
                return (255, 255, 255), (0, 0, 0)

        # pytesseract.image_to_data 获取文字框；若未安装 Tesseract 则自动回退 Windows Media OCR
        try:
            data = pytesseract.image_to_data(image, lang='chi_sim', output_type=pytesseract.Output.DICT)
        except Exception as e:
            if os.path.exists(input_path):
                data = self._ocr_boxes_via_windows_media(input_path)
            else:
                raise RuntimeError(f"OCR 提取失败，请检查是否安装了 Tesseract 及 chi_sim 包: {e}")

        # 若检测到横版图片为 90°/270° 旋转的扫描件，先旋转至正向完成原位覆盖后再保存正方向图片
        rot_angle = data.get('rotation', 0) if isinstance(data, dict) else 0
        if rot_angle in (90, 270):
            image = image.rotate(rot_angle, expand=True)

        orig_img_copy = image.copy()
        img_w, img_h = image.size
        draw = ImageDraw.Draw(image)

        def get_font_for_height(h):
            font_size = max(12, int(round(h * 0.98)))
            for font_name in ["msyh.ttc", "simhei.ttf", "simsun.ttc"]:
                try:
                    return ImageFont.truetype(font_name, font_size)
                except Exception:
                    pass
            try:
                return ImageFont.load_default()
            except Exception:
                return None

        orig_lines = []
        conv_lines = []
        if data.get('lines'):
            lines = data['lines']
            total_words = max(1, sum(len(l.get('words', [])) for l in lines))
            current = 0
            for line in lines:
                orig_line = line.get('text', '')
                words = line.get('words', [])
                if not orig_line or not words:
                    continue
                orig_lines.append(orig_line)
                conv_line = self.engine.convert_text(orig_line)
                conv_lines.append(conv_line)
                same_len = len(conv_line) == len(words)

                cjk_hs = [w.get('height', 10) for w in words if w.get('text') and '\u4e00' <= w.get('text')[0] <= '\u9fff']
                if not cjk_hs:
                    cjk_hs = [w.get('height', 10) for w in words if w.get('height', 0) > 0]
                cjk_hs.sort()
                line_h = cjk_hs[min(len(cjk_hs) - 1, int(len(cjk_hs) * 0.75))] if cjk_hs else 16
                font = get_font_for_height(line_h)

                changed_items = []
                for w_idx, w_info in enumerate(words):
                    char_text = w_info.get('text', '')
                    if same_len:
                        conv_char = conv_line[w_idx]
                    elif w_idx < len(conv_line):
                        conv_char = conv_line[w_idx]
                    else:
                        conv_char = self.engine.convert_text(char_text)

                    # 未发生简繁变化的文字 100% 保留原图清晰字迹，避免擦除与重绘
                    if char_text == conv_char:
                        current += 1
                        continue

                    x = int(w_info.get('left', 0))
                    y = int(w_info.get('top', 0))
                    w = int(w_info.get('width', 10))
                    h = int(w_info.get('height', 10))

                    bg_col, fg_col = sample_colors(orig_img_copy, x, y, w, h)
                    changed_items.append((x, y, w, h, conv_char, bg_col, fg_col))
                    current += 1

                # 第一遍：先对整行所有发生变动的字统一进行背景采样抹除（外扩 1px 消除旧字残影，避免后续字符遮挡前一个字符）
                for x, y, w, h, conv_char, bg_col, fg_col in changed_items:
                    draw.rectangle([max(0, x - 1), max(0, y - 1), min(img_w, x + w + 1), min(img_h, y + h + 1)], fill=bg_col)

                # 第二遍：将目标繁体字根据字形 bbox 居中精准绘制到原字框内
                for x, y, w, h, conv_char, bg_col, fg_col in changed_items:
                    try:
                        tb = draw.textbbox((0, 0), conv_char, font=font)
                        if isinstance(tb, (list, tuple)) and len(tb) >= 4:
                            tb_w = tb[2] - tb[0]
                            tb_h = tb[3] - tb[1]
                            cx = x + w / 2.0
                            cy = y + h / 2.0
                            dx = int(round(cx - tb_w / 2.0 - tb[0]))
                            dy = int(round(cy - tb_h / 2.0 - tb[1]))
                        else:
                            dx = x
                            dy = y - int(h * 0.08)
                    except Exception:
                        dx = x
                        dy = y - int(h * 0.08)
                    draw.text((dx, dy), conv_char, fill=fg_col, font=font)

                if progress_cb:
                    progress_cb(int((current / total_words) * 100))
        else:
            n_boxes = len(data.get('level', []))
            for i in range(n_boxes):
                text = data['text'][i].strip() if data.get('text') and i < len(data['text']) else ""
                if text:
                    orig_lines.append(text)
                    conv_text = self.engine.convert_text(text)
                    conv_lines.append(conv_text)
                    if text != conv_text:
                        x = data['left'][i]
                        y = data['top'][i]
                        w = data['width'][i]
                        h = data['height'][i]
                        font = get_font_for_height(h)
                        bg_col, fg_col = sample_colors(orig_img_copy, x, y, w, h)
                        draw.rectangle([max(0, x - 1), max(0, y - 1), min(img_w, x + w + 1), min(img_h, y + h + 1)], fill=bg_col)
                        try:
                            tb = draw.textbbox((0, 0), conv_text, font=font)
                            if isinstance(tb, (list, tuple)) and len(tb) >= 4:
                                tb_w = tb[2] - tb[0]
                                tb_h = tb[3] - tb[1]
                                cx = x + w / 2.0
                                cy = y + h / 2.0
                                dx = int(round(cx - tb_w / 2.0 - tb[0]))
                                dy = int(round(cy - tb_h / 2.0 - tb[1]))
                            else:
                                dx = x
                                dy = y - int(h * 0.08)
                        except Exception:
                            dx = x
                            dy = y - int(h * 0.08)
                        draw.text((dx, dy), conv_text, fill=fg_col, font=font)

                if progress_cb and n_boxes > 0:
                    progress_cb(int(((i + 1) / n_boxes) * 100))

        image.save(output_path)
        # 保存无损精确繁体 OCR 文本缓存 sidecar 文件及原文缓存文件
        try:
            with open(output_path + ".ocr.txt", "w", encoding="utf-8") as f:
                f.write("\n".join(conv_lines))
            with open(output_path + ".orig_ocr.txt", "w", encoding="utf-8") as f_orig:
                f_orig.write("\n".join(orig_lines))
        except Exception:
            pass
        if progress_cb:
            progress_cb(100)