import docx

def convert_paragraph_runs_safely(paragraph, converter_func):
    """
    通过字符偏移映射，解决中文词语被拆分在多个 run 节点的问题
    同时完全保留每个 run 的原有独立样式。
    """
    runs = paragraph.runs
    if not runs:
        return

    # 1. 拼接文本并记录各个 run 的切片边界
    run_slices = []
    full_chars = []
    current_idx = 0

    for r in runs:
        text = r.text
        length = len(text)
        run_slices.append((current_idx, current_idx + length))
        current_idx += length
        full_chars.append(text)

    full_text = "".join(full_chars)
    if not full_text:
        return

    # 2. 全文送入转换引擎（此时分词器与模型拥有完整上下文）
    converted_full = converter_func(full_text)

    # 3. 校验字符长度一致性并分段回填
    # 简繁转换中大部分为等长替换，若出现不等长情况需进行安全兜底
    if len(converted_full) == len(full_text):
        for r, (start, end) in zip(runs, run_slices):
            r.text = converted_full[start:end]
    else:
        # 异常兜底：若有特殊替换导致长度不等，将全文赋给首个 run，清空后续 run
        runs[0].text = converted_full
        for r in runs[1:]:
            r.text = ""