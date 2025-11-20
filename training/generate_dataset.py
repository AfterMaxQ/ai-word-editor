# training/generate_dataset.py
import json
import yaml
import os

# --- 配置路径 ---
# 脚本期望在 training/ 目录下运行
# ../prompts/system_prompt.txt 指向项目根目录下的 prompts 文件夹
SYSTEM_PROMPT_PATH = os.path.join("..", "prompts", "system_prompt.txt")
TRAINING_DATA_YAML_PATH = "training_data.yaml"
# ../dataset.jsonl 指向项目根目录
OUTPUT_JSONL_PATH = os.path.join("..", "dataset.jsonl")


def main():
    """
    读取 system_prompt 和 training_data.yaml，
    并生成 Ollama 微调所需的 dataset.jsonl 文件。
    """
    print("--- 开始生成微调数据集 ---")

    # 1. 读取 System Prompt
    try:
        with open(SYSTEM_PROMPT_PATH, 'r', encoding='utf-8') as f:
            system_prompt_content = f.read()
        print(f"✅ 成功读取系统提示词: {SYSTEM_PROMPT_PATH}")
    except FileNotFoundError:
        print(f"❌ 错误: 找不到系统提示词文件 at '{SYSTEM_PROMPT_PATH}'")
        print("请确保你在 'training/' 目录下运行此脚本，并且 'prompts/system_prompt.txt' 存在。")
        return

    # 2. 读取 YAML 训练数据
    try:
        with open(TRAINING_DATA_YAML_PATH, 'r', encoding='utf-8') as f:
            training_data = yaml.safe_load(f)
        if not isinstance(training_data, list):
            print(f"❌ 错误: '{TRAINING_DATA_YAML_PATH}' 的内容不是一个有效的 YAML 列表。")
            return
        print(f"✅ 成功读取 {len(training_data)} 个训练样本: {TRAINING_DATA_YAML_PATH}")
    except FileNotFoundError:
        print(f"❌ 错误: 找不到训练数据文件 at '{TRAINING_DATA_YAML_PATH}'")
        return
    except yaml.YAMLError as e:
        print(f"❌ 错误: 解析 YAML 文件时出错: {e}")
        return

    # 3. 生成并写入 dataset.jsonl
    generated_count = 0
    with open(OUTPUT_JSONL_PATH, 'w', encoding='utf-8') as f_out:
        for i, sample in enumerate(training_data):
            user_prompt = sample.get('user_prompt')
            assistant_response_str = sample.get('assistant_response')

            if not user_prompt or not assistant_response_str:
                print(f"⚠️ 警告: 跳过第 {i + 1} 个样本，因为它缺少 'user_prompt' 或 'assistant_response'。")
                continue

            # 验证 assistant_response 是否是有效的 JSON
            try:
                # 去掉 YAML 多行字符串可能带来的首尾空白
                assistant_response_str = assistant_response_str.strip()
                json.loads(assistant_response_str)
            except json.JSONDecodeError as e:
                print(f"❌ 错误: 第 {i + 1} 个样本的 'assistant_response' 不是有效的 JSON。跳过。")
                print(f"   错误详情: {e}")
                print(f"   问题内容: {assistant_response_str[:100]}...")
                continue

            # 构建 Ollama 格式的 JSON 对象
            ollama_record = {
                "messages": [
                    {
                        "role": "system",
                        "content": system_prompt_content
                    },
                    {
                        "role": "user",
                        "content": user_prompt.strip()
                    },
                    {
                        "role": "assistant",
                        "content": assistant_response_str
                    }
                ]
            }

            # 将该对象作为一行写入 .jsonl 文件
            f_out.write(json.dumps(ollama_record, ensure_ascii=False) + '\n')
            generated_count += 1

    print(f"✅ 成功生成 {generated_count} 条记录到: {OUTPUT_JSONL_PATH}")
    print("--- 数据集生成完毕 ---")


if __name__ == "__main__":
    # 确保当前工作目录是脚本所在的目录
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    main()