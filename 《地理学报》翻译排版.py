import re
import hashlib
import requests
import time
import sys
from docx import Document

# ==================== 有道智云配置 ====================
# 请将以下变量替换为您的实际应用信息
YOUDAO_APP_ID = "778da5d057c333c0"          # 应用ID
YOUDAO_APP_SECRET = "EGmD0yrPudA78FWjVRz1lVtzSkllclrE"  # 应用密钥
REGION = ""                 # 区域（根据文档填写，如不需要可留空）

# 有道翻译API端点
YOUDAO_API_URL = "https://openapi.youdao.com/api"

# 请求间隔（秒），避免限流
REQUEST_INTERVAL = 2
MAX_RETRIES = 3  # 最大重试次数

# ==================== 翻译函数 ====================
def translate_with_youdao(text, from_lang='zh-CHS', to_lang='en', retry=0):
    """
    调用有道翻译API进行文本翻译，包含重试机制
    """
    if not text.strip():
        return text

    salt = str(int(time.time() * 1000))
    sign_str = YOUDAO_APP_ID + text + salt + YOUDAO_APP_SECRET
    sign = hashlib.md5(sign_str.encode('utf-8')).hexdigest()

    params = {
        'q': text,
        'from': from_lang,
        'to': to_lang,
        'appKey': YOUDAO_APP_ID,
        'salt': salt,
        'sign': sign
    }
    if REGION:
        params['region'] = REGION

    try:
        response = requests.post(YOUDAO_API_URL, data=params)
        result = response.json()
        if result.get('errorCode') == '0':
            translation = result.get('translation', [''])[0]
            return translation
        else:
            error_code = result.get('errorCode')
            print(f"翻译失败，错误码：{error_code}，信息：{result.get('errorMsg', '')}")
            if error_code == '411' and retry < MAX_RETRIES:
                # 频率限制，等待更长时间后重试
                wait_time = REQUEST_INTERVAL * (2 ** retry)  # 指数退避
                print(f"频率限制，{wait_time}秒后重试...")
                time.sleep(wait_time)
                return translate_with_youdao(text, from_lang, to_lang, retry+1)
            else:
                return text
    except Exception as e:
        print(f"请求异常：{e}")
        if retry < MAX_RETRIES:
            wait_time = REQUEST_INTERVAL * (2 ** retry)
            print(f"{wait_time}秒后重试...")
            time.sleep(wait_time)
            return translate_with_youdao(text, from_lang, to_lang, retry+1)
        else:
            return text

# ==================== 文档处理函数 ====================
def extract_ref_text(docx_path):
    """
    从Word文档中提取“参考文献(References)”之后的所有文本
    """
    doc = Document(docx_path)
    ref_started = False
    ref_lines = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not ref_started and "参考文献(References)" in text:
            ref_started = True
            idx = text.find("参考文献(References)")
            if idx != -1:
                after_text = text[idx + len("参考文献(References)"):].strip()
                if after_text:
                    ref_lines.append(after_text)
            continue
        if ref_started and text:
            ref_lines.append(text)

    return "\n".join(ref_lines)

def split_ref_entries(text):
    """
    将参考文献文本按编号分割为单个条目
    """
    pattern = r'(\[\d+\])'
    parts = re.split(pattern, text)
    entries = []
    for i in range(1, len(parts), 2):
        idx = parts[i]
        content = parts[i+1] if i+1 < len(parts) else ''
        entry = idx + content.strip()
        if entry:
            entries.append(entry)
    return entries

def is_chinese_ref(entry):
    """判断是否包含中文字符"""
    return bool(re.search(r'[\u4e00-\u9fa5]', entry))

def process_document(docx_path):
    print("正在提取参考文献...")
    ref_text = extract_ref_text(docx_path)
    if not ref_text:
        print("未找到参考文献部分，请检查文档中是否包含“参考文献(References)”字样。")
        return

    print("分割参考文献条目...")
    entries = split_ref_entries(ref_text)
    print(f"共发现 {len(entries)} 条文献。")

    results = []
    for i, entry in enumerate(entries, 1):
        print(f"处理第 {i} 条...")
        if is_chinese_ref(entry):
            # 提取编号
            num_match = re.match(r'^(\[\d+\])', entry)
            if num_match:
                num = num_match.group(1)
                content = entry[len(num):].strip()
            else:
                num = ''
                content = entry

            # 翻译整个内容（去掉编号后）
            translated = translate_with_youdao(content)
            # 构建新条目：编号 + 翻译内容 + [原始中文]
            new_entry = f" {num} {translated} [{content}]"
            results.append(new_entry)
            # 等待避免限流
            time.sleep(REQUEST_INTERVAL)
        else:
            # 英文文献直接保留
            results.append(entry)

    # 输出结果
    output_text = "\n".join(results)
    print("\n处理完成，结果如下：\n")
    print(output_text)

    # 保存到文件
    output_file = "processed_references_simple.txt"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(output_text)
    print(f"\n结果已保存至 {output_file}")

# ==================== 主程序 ====================
if __name__ == "__main__":
    default_path = r"C:\Users\ZJZ\Desktop\社会地理论文\（二稿）城市绿色绅士化与生态系统服务的互动关系及形成机制—以长沙市主城区为例.docx"
    if len(sys.argv) > 1:
        docx_path = sys.argv[1]
    else:
        docx_path = default_path
    process_document(docx_path)