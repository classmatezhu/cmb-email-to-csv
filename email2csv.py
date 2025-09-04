import email
import os
import sys
import csv
from bs4 import BeautifulSoup, Tag

def extract_html_from_email_file(file_path):
    """
    从保存为 .txt 的邮件原文文件中提取 HTML 内容。
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            raw_email = f.read()
    except FileNotFoundError:
        print(f"错误：文件 '{file_path}' 未找到。")
        return None
    except Exception as e:
        print(f"读取文件时发生错误: {e}")
        return None

    msg = email.message_from_string(raw_email)

    for part in msg.walk():
        if part.get_content_type() == 'text/html':
            payload = part.get_payload(decode=True)
            charset = part.get_content_charset() or 'utf-8'
            try:
                html_content = payload.decode(charset)
                return html_content
            except (UnicodeDecodeError, AttributeError):
                if isinstance(payload, str):
                    return payload
                print(f"使用字符集 '{charset}' 解码失败，请检查邮件编码。")
                return None
    return None

def extract_rows_from_html(html_text):
    """
    按文档流顺序从HTML文本中提取需要写入 csv 的行。
    """
    soup = BeautifulSoup(html_text, 'html.parser')
    rows = []

    for node in soup.descendants:
        if not isinstance(node, Tag):
            continue
        if node.name == 'strong':
            rows.append([node.get_text(strip=True)])
        elif node.name == 'span' and node.get('id') == 'fixBand15':
            divs = node.find_all('div')
            row = [div.get_text(strip=True) for div in divs]
            rows.append(row)
    return rows

def write_to_csv(rows, csv_path):
    """
    用 UTF-8 with BOM 写入 csv。
    """
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerows(rows)
    print(f"✅ 成功生成 CSV 文件: {csv_path} (编码: UTF-8 with BOM)")

def main():
    """
    主函数，协调整个流程。
    """
    # 检查命令行参数
    if len(sys.argv) != 2:
        print("❌ 使用方法错误！")
        print(f"正确用法: python {sys.argv[0]} <邮件文件路径>")
        print(f"示例: python {sys.argv[0]} email.txt")
        sys.exit(1)
    
    email_file = sys.argv[1]
    
    # 生成输出文件名（基于输入文件名）
    base_name = os.path.splitext(email_file)[0]
    csv_file = f"{base_name}_output.csv"

    if not os.path.exists(email_file):
        print(f"❌ 错误：输入文件 '{email_file}' 不存在。")
        sys.exit(1)

    print(f"正在从 '{email_file}' 提取 HTML 内容...")
    html_content = extract_html_from_email_file(email_file)

    if not html_content:
        print("❌ 未在邮件中找到 HTML 内容。")
        sys.exit(1)
  
    print("提取 HTML 成功，正在解析数据...")
    rows = extract_rows_from_html(html_content)

    if not rows:
        print("⚠️ 在 HTML 内容中没有找到可供提取的数据。")
        sys.exit(1)

    print(f"解析完成，正在将数据写入 '{csv_file}'...")
    write_to_csv(rows, csv_file)

if __name__ == "__main__":
    main()
