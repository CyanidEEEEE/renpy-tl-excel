import os
import re
import concurrent.futures
import datetime
import traceback
import multiprocessing
import pandas as pd

def timestamp():
    # 返回当前时间戳字符串
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")

def normalize_lang(lang):
    """
    归一化语言名称：
    如果语言以 "c" 开头，则去掉首字母 "c"（例如将 "cchinese" 变为 "chinese"），
    否则直接返回原始字符串。
    """
    return lang[1:] if lang.startswith("c") else lang

def process_file(filepath, language, all_translations):
    """处理单个 .rpy 文件，应用翻译。"""
    print(f"[{timestamp()}] 正在处理文件：{filepath}")

    try:
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
    except FileNotFoundError:
        print(f"[{timestamp()}] 错误：未找到文件 '{filepath}'。跳过该文件。")
        return
    except Exception as e:
        print(f"[{timestamp()}] 读取文件 '{filepath}' 时出错：{type(e).__name__}: {e}\n{traceback.format_exc()}")
        return

    file_modified = False
    current_block = None  # 当前翻译块信息（包含语言和标识符）
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        # 检查翻译块头，例如 "translate chinese startday4_881cedf5:"
        translate_match = re.match(r'^\s*translate\s+(\w+)\s+(\w+):\s*$', line)
        if translate_match:
            current_block = {
                'language': translate_match.group(1),
                'identifier': translate_match.group(2)
            }
            new_lines.append(line)
            i += 1
            continue

        # 如果处于翻译块内且当前行不是注释，则尝试匹配对话行
        if current_block is not None and not line.lstrip().startswith("#"):
            # 根据行首是否以双引号开头判断是否有 actor
            if line.lstrip().startswith('"'):
                # 无 actor 的情况
                dialogue_match = re.match(r'^(\s*)"(.*)"\s*$', line)
                if dialogue_match:
                    indent = dialogue_match.group(1)
                    actor = ""
                    dialogue_text = dialogue_match.group(2)
                else:
                    dialogue_text = None
            else:
                # 有 actor 的情况
                dialogue_match = re.match(r'^(\s*)(\S+)\s*"(.*)"\s*$', line)
                if dialogue_match:
                    indent = dialogue_match.group(1)
                    actor = dialogue_match.group(2)
                    dialogue_text = dialogue_match.group(3)
                else:
                    dialogue_text = None

            if dialogue_text is not None:
                # 如果文本中包含字面 "\n"，则认为 "\n" 后面的部分为当前翻译；否则整个文本为当前翻译
                if "\\n" in dialogue_text:
                    current_candidate = dialogue_text.split("\\n", 1)[1]
                else:
                    current_candidate = dialogue_text

                found_translation = None
                # 遍历 Excel 中加载的翻译记录，寻找与当前翻译块匹配的记录
                for translation in all_translations:
                    if (translation['identifier'] and
                        normalize_lang(current_block['language']) == normalize_lang(language) and
                        current_block['identifier'] == translation['identifier']):
                        found_translation = translation['translated_text']
                        break

                # 如果找到匹配的翻译且当前文件中的翻译与 Excel 中的不一致，则更新
                if found_translation is not None and found_translation != current_candidate:
                    if actor:
                        new_line = f'{indent}{actor}"{found_translation}"\n'
                        print(f"[{timestamp()}] 使用标识符 '{current_block['identifier']}' 替换文件 {filepath} 中 actor '{actor}' 的对话。")
                    else:
                        new_line = f'{indent}"{found_translation}"\n'
                        print(f"[{timestamp()}] 使用标识符 '{current_block['identifier']}' 替换文件 {filepath} 中无 actor 的对话。")
                    new_lines.append(new_line)
                    file_modified = True
                    # 替换后退出当前翻译块
                    current_block = None
                    i += 1
                    continue
                else:
                    # 若匹配成功且内容相同，则不修改，退出当前翻译块
                    new_lines.append(line)
                    current_block = None
                    i += 1
                    continue

        # 处理 old/new 结构
        old_match = re.match(r'^(\s*)old\s*"(.*?)"\s*(#.*)?$', line)
        if old_match:
            original_text_in_file = old_match.group(2)
            print(f"[{timestamp()}] 在文件 {filepath} 中找到 'old' 行：原文为 '{original_text_in_file}'")
            new_lines.append(line)
            if i + 1 < len(lines):
                new_line_candidate = lines[i+1]
                new_match = re.match(r'^(\s*)new\s*"(.*?)"(.*)$', new_line_candidate)
                if new_match:
                    replaced = False
                    for translation in all_translations:
                        if (not translation['identifier'] and 
                            translation['original_text'] == original_text_in_file):
                            indent_new = new_match.group(1)
                            new_translation_text = translation['translated_text']
                            new_line_replaced = f'{indent_new}new "{new_translation_text}"\n'
                            print(f"[{timestamp()}] 用 '{new_translation_text}' 替换文件 {filepath} 中的 'new' 行。")
                            new_lines.append(new_line_replaced)
                            file_modified = True
                            replaced = True
                            break
                    if replaced:
                        i += 2
                        continue
                    else:
                        new_lines.append(new_line_candidate)
                        i += 2
                        continue
            i += 1
            continue

        # 其他行直接保留
        new_lines.append(line)
        i += 1

    if file_modified:
        try:
            with open(filepath, 'w', encoding='utf-8-sig') as f:
                f.writelines(new_lines)
            print(f"[{timestamp()}] 文件已更新：{filepath}")
        except Exception as e:
            print(f"[{timestamp()}] 写入文件 '{filepath}' 时出错：{type(e).__name__}: {e}\n{traceback.format_exc()}")

def update_rpy_translations(language):
    """主函数：更新指定语言的翻译。"""
    print(f"[{timestamp()}] 开始更新语言 {language} 的翻译。")

    excel_file = os.path.join(os.getcwd(), f"{language}.xlsx")
    rpy_dir = os.path.join(os.getcwd(), "game", "tl", language)

    if not os.path.isdir(rpy_dir):
        print(f"[{timestamp()}] 错误：目录 '{rpy_dir}' 不存在。")
        return

    all_translations = []
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        print(f"[{timestamp()}] 加载 Excel 文件：{excel_file}")

        for index, row in df.iterrows():
            prefix = row[0] if not pd.isna(row[0]) else ""
            original_text = row[1] if not pd.isna(row[1]) else ""
            translated_text = row[2] if not pd.isna(row[2]) else ""
            location = row[4] if not pd.isna(row[4]) else ""
            identifier = row[5] if not pd.isna(row[5]) else ""

            if not original_text or not translated_text or not location:
                print(f"[{timestamp()}] 警告：跳过第 {index + 2} 行。")
                continue

            all_translations.append({
                'prefix': prefix,
                'original_text': original_text,
                'translated_text': translated_text,
                'location': location,
                'identifier': identifier,
            })

    except FileNotFoundError:
        print(f"[{timestamp()}] 错误：找不到 Excel 文件 '{excel_file}'。")
        return
    except Exception as e:
        print(f"[{timestamp()}] 读取 Excel 文件时出错：{type(e).__name__}: {e}\n{traceback.format_exc()}")
        return

    print(f"[{timestamp()}] 处理 Excel 数据完成。加载了 {len(all_translations)} 条翻译。")

    rpy_files = [os.path.join(rpy_dir, f) for f in os.listdir(rpy_dir) if f.endswith(".rpy")]
    print(f"[{timestamp()}] 找到 {len(rpy_files)} 个 .rpy 文件。")

    with concurrent.futures.ProcessPoolExecutor(max_workers=multiprocessing.cpu_count() * 2) as executor:
        futures = [executor.submit(process_file, filepath, language, all_translations) for filepath in rpy_files]
        print(f"[{timestamp()}] 已将所有文件提交到进程池。")

        for future in concurrent.futures.as_completed(futures):
            try:
                future.result()
            except Exception as e:
                print(f"[{timestamp()}] 处理文件时发生错误：{type(e).__name__}: {e}\n{traceback.format_exc()}")

    print(f"[{timestamp()}] 语言 {language} 的翻译更新完成。")

if __name__ == "__main__":
    language = input("请输入目标语言代码（例如：cchinese）：")
    update_rpy_translations(language)
    print("程序结束。")
