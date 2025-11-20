# finetune.py
import os

os.environ['HTTP_PROXY'] = 'http://127.0.0.1:7897'
os.environ['HTTPS_PROXY'] = 'http://127.0.0.1:7897'

import torch
from datasets import load_dataset
from transformers import AutoTokenizer, AutoModelForCausalLM, BitsAndBytesConfig, TrainingArguments
from peft import LoraConfig, get_peft_model
from trl import SFTTrainer


# --- 新增：数据格式化函数 ---
# 这个函数的作用是接收一个数据样本（包含 "messages" 字段）
# 然后使用分词器的聊天模板，将其转换为一个格式化的字符串
def format_dataset(example, tokenizer):
    # apply_chat_template 会根据模型预设的格式，自动添加角色名和特殊token
    # 例如，它会把 [{"role": "user", "content": "Hi"}] 转换成类似 "<|im_start|>user\nHi<|im_end|>" 的格式
    # tokenize=False 表示我们只想得到格式化后的字符串，而不是直接分词后的ID
    formatted_prompt = tokenizer.apply_chat_template(
        example["messages"],
        tokenize=False
    )
    # 我们返回一个新的字典，其中 "text" 键对应的值就是我们格式化好的字符串
    return {"text": formatted_prompt}


# --- 函数结束 ---


def main():
    # --- 1. 定义模型和数据集的路径 ---
    model_id = "Qwen/Qwen2.5-7B-Instruct"
    dataset_path = "dataset.jsonl"

    print(f"--- 步骤 1: 加载数据集 {dataset_path} ---")
    dataset = load_dataset("json", data_files=dataset_path, split="train")
    print("\n数据集加载成功！")

    # --- 2. 配置模型加载方式 (4-bit 量化) ---
    print(f"\n--- 步骤 2: 配置并加载模型 {model_id} ---")
    bnb_config = BitsAndBytesConfig(
        load_in_4bit=True,
        bnb_4bit_quant_type="nf4",
        bnb_4bit_compute_dtype=torch.bfloat16,
        bnb_4bit_use_double_quant=False,
    )

    # --- 3. 加载分词器 (Tokenizer) 和模型 ---
    print("\n正在加载分词器...")
    tokenizer = AutoTokenizer.from_pretrained(model_id)
    tokenizer.pad_token = tokenizer.eos_token

    print("\n正在加载模型 (这可能需要一些时间来下载)...")
    model = AutoModelForCausalLM.from_pretrained(
        model_id,
        quantization_config=bnb_config,
        device_map="auto"
    )
    print("\n模型和分词器加载成功！")

    # --- 新增步骤：格式化数据集 ---
    print("\n--- 正在使用聊天模板格式化数据集 ---")
    # 使用 .map() 方法将我们的格式化函数应用到数据集的每一个样本上
    # 我们需要把 tokenizer 传给函数，所以使用 fn_kwargs 参数
    formatted_dataset = dataset.map(format_dataset, fn_kwargs={'tokenizer': tokenizer})

    # 打印一个格式化后的样本，看看效果
    print("\n格式化完成！看一个格式化后的样本:")
    print(formatted_dataset[0]['text'])
    # --- 格式化结束 ---

    # --- 步骤 4: 配置 LoRA (PEFT) ---
    print("\n--- 步骤 4: 配置 LoRA (低秩适应) ---")
    lora_config = LoraConfig(
        r=8,
        lora_alpha=16,
        target_modules=["q_proj", "k_proj", "v_proj", "o_proj", "gate_proj", "up_proj", "down_proj"],
        lora_dropout=0.05,
        bias="none",
        task_type="CAUSAL_LM",
    )
    model = get_peft_model(model, lora_config)
    model.print_trainable_parameters()

    # --- 步骤 5: 配置训练参数 ---
    print("\n--- 步骤 5: 配置训练参数 ---")
    training_args = TrainingArguments(
        output_dir="./qwen2-finetuned",
        per_device_train_batch_size=1,
        gradient_accumulation_steps=4,
        learning_rate=2e-4,
        num_train_epochs=1,
        logging_steps=10,
        save_steps=50,
        fp16=True,
    )

    # --- 步骤 6: 初始化训练器并开始训练 ---
    print("\n--- 步骤 6: 初始化训练器 ---")
    trainer = SFTTrainer(
        model=model,
        # *** 修改点 1: 使用格式化后的数据集 ***
        train_dataset=formatted_dataset,
        peft_config=lora_config,
        # *** 修改点 2: 指定新的文本字段名 ***
        dataset_text_field="text",
        max_seq_length=1024,
        tokenizer=tokenizer,
        args=training_args,
    )

    print("\n--- 开始训练！ ---")
    trainer.train()
    print("\n--- 训练完成！ ---")

    # --- 步骤 7: 保存最终的模型 ---
    print("\n--- 步骤 7: 保存最终的LoRA适配器 ---")
    final_model_path = "./qwen2-finetuned-final"
    trainer.save_model(final_model_path)
    print(f"模型已保存到: {final_model_path}")


if __name__ == "__main__":
    main()