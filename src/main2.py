import gradio as gr
import openai
from llm import LLM  # 导入语言模型类，可能用于生成报告内容
from config import Config  # 导入配置管理类
import argparse
from input_parser import parse_input_text
from ppt_generator import generate_presentation
from template_manager import load_template, get_layout_mapping, print_layouts
from layout_manager import LayoutManager
from logger import LOG  # 引入 LOG 模块

# 定义提示词文件路径
PROMPT_FILE_PATH = "prompts/formatter.txt"

# 读取本地文件中的提示词
def read_prompts(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        prompts = f.read()
    return prompts

# 使用 OpenAI GPT-4 模型处理用户输入
def query_gpt(user_input,config):
    
    llm = LLM(config)  # 创建语言模型实例
    # 读取提示词
    prompts = read_prompts(PROMPT_FILE_PATH)

    report = llm.generate_report(prompts,user_input)
    return report


# 处理用户输入，生成 Markdown 内容
def generate_ppt(user_input):
    config = Config()  # 创建配置实例
    # 使用模型生成内容
    markdown_content = query_gpt(user_input,config)

    # 打印到控制台
    print("Generated Markdown Content:\n", markdown_content)  # 打印生成的内容

    # 生成的 Markdown 文件名
    output_file = "outputs/output.md"
    

    # 加载 PowerPoint 模板，并打印模板中的可用布局
    prs = load_template(config.ppt_template)  # 加载模板文件
    LOG.info("可用的幻灯片布局:")  # 记录信息日志，打印可用布局
    print_layouts(prs)  # 打印模板中的布局

    # 初始化 LayoutManager，使用配置文件中的 layout_mapping
    layout_manager = LayoutManager(config.layout_mapping)

    # 调用 parse_input_text 函数，解析输入文本，生成 PowerPoint 数据结构
    powerpoint_data, presentation_title = parse_input_text(markdown_content, layout_manager)

    LOG.info(f"解析转换后的 ChatPPT PowerPoint 数据结构:\n{powerpoint_data}")  # 记录调试日志，打印解析后的 PowerPoint 数据

    # 定义输出 PowerPoint 文件的路径
    output_pptx = f"outputs/{presentation_title}.pptx"
    
    # 调用 generate_presentation 函数生成 PowerPoint 演示文稿
    generate_presentation(powerpoint_data, config.ppt_template, output_pptx)

    return output_pptx

# Gradio 界面定义
def main():
    with gr.Blocks() as demo:
        gr.Markdown("## PPT 文件生成器")
        
        user_input = gr.Textbox(label="请输入内容", placeholder="这里输入您的内容...")
        
        submit_btn = gr.Button("生成PPT文件")
        
        output = gr.File(label="生成的PPT文件")
        
        submit_btn.click(fn=generate_ppt, inputs=user_input, outputs=output)

    demo.launch()

if __name__ == "__main__":
    main()