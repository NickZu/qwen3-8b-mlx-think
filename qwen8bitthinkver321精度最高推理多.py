#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
业主大会投票核验系统（终极版 - 含评分）
- 多文件择优：正常接通无意愿（3）优于异常通话（4）
- 增强JSON提取
- 明确提示词（带示例 + 禁止思考过程）
- 长文本支持（max_tokens=1000）
- 身份模糊匹配（拼音+身份证后四位）
- 调试日志：可配置路径，包含票号和文件名
- 新增投票评分（0~100），辅助人工复核
- 安全处理：防止 identity/vote_intent 为 None

适配 mlx-lm >= 0.31.1
"""

import os
import json
import re
import pandas as pd
from tqdm import tqdm
from datetime import datetime

# 模糊匹配所需库
try:
    from fuzzywuzzy import fuzz
    from pypinyin import lazy_pinyin
except ImportError:
    print("请安装依赖库：pip install fuzzywuzzy pypinyin python-Levenshtein")
    exit(1)

# mlx-lm 相关导入
from mlx_lm import load, generate
from mlx_lm.sample_utils import make_sampler

# ====================== 【用户配置区】一键修改 ======================
EXCEL_PATH = "/Users/nickzu/Downloads/2.xlsx"   
TXT_FOLDER = "/Users/nickzu/Desktop/识477" 
MODEL_NAME = "/Users/nickzu/Qwen/Qwen3-8B-MLX-4bit"
RESULT_SAVE_PATH = "/Users/nickzu/Desktop/投票结果_477.xlsx"
LOG_FILE_PATH = "/Users/nickzu/Desktop/failed_responses.log"   # 可自定义日志路径
EXCEL_KEY_COLUMNS = {
    "票号": "票权编号",
    "业主姓名": "姓名",
    "身份证号": "证件号",
    "投票内容": "其他备注"          # Excel中的N列，人工备注
}
# =====================================================================

# 加载MLX模型
print("正在加载模型...")
model, tokenizer = load(MODEL_NAME)

# ====================== 【系统提示词（带示例，强制JSON输出，禁止思考过程）】======================
SYSTEM_PROMPT = """
你是一个业主大会投票核验专家。请根据录音文本，严格按照以下规则提取信息，并**只输出一个JSON对象**，不要有任何其他文字或解释。**禁止输出任何思考过程、解释或Markdown标记，直接输出JSON。**

【规则】
1. 通话状态：
   - 正常接通：有双方对话，并且存在实质性交流（如听取介绍、表达意见、核对身份等）。即使开头有抱怨、忙碌，只要最终完成了投票交流，仍属正常接通。
   - 异常状态：无人接听、忙音、空文本、纯杂音、语音提示、接通后仅说“在忙”就挂断且无任何投票相关内容。
2. 身份核验：从文本中提取业主姓名、身份证后四位、自称身份（业主本人/亲属/未明确）。注意语音转写可能有错别字，请合理推断。
3. 投票意愿（重要）：
   - 候选人选择（option1）：从对话中追踪业主**最终确认**的选择。业主可能多次更改，以最后一次明确的表态为准。
     * 例如：“排除1号”、“只选10号”、“选1号和10号”、“弃权”、“选1-9号”等。
     * 必须忽略试探性或模糊的中间表达，只记录最终决定。
   - 草案表态（option2）：同样取业主最终确认的表态（“赞同”/“反对”/“弃权”）。
   - full_text：完整描述投票意愿（候选人和草案的最终结果）。
4. 推理说明（explanation）：简要描述关键步骤，特别是如果业主中途改变主意，要说明“先说要选X，后改为选Y，最终确定为Z”。如果全程一致，正常描述即可。若通话开始有抱怨但后续完成投票，也要说明“业主起初抱怨，后完成投票”。

【输出示例1】业主最终选择1号和10号，草案赞同，且前期有过变化：
{
  "call_status": "正常接通",
  "identity": {
    "name": "陈永军",
    "id_last4": "0010",
    "self_claim": "业主本人"
  },
  "vote_intent": {
    "option1": "选1号和10号",
    "option2": "赞同",
    "full_text": "最终选择1号和10号候选人，草案赞同"
  },
  "explanation": "业主先表示要选10号，后改口说选1号和10号，确认最终选择；草案明确表示同意。"
}

【输出示例2】业主最终弃权，草案赞同，且开头有抱怨：
{
  "call_status": "正常接通",
  "identity": {
    "name": "李刚毅",
    "id_last4": "3779",
    "self_claim": "业主本人"
  },
  "vote_intent": {
    "option1": "弃权",
    "option2": "赞同",
    "full_text": "候选人弃权，草案赞同"
  },
  "explanation": "业主起初抱怨电话多，但仍完成投票，最终选择弃权，草案赞同。"
}

【输出示例3】通话仅为短暂交流，无投票意愿：
{
  "call_status": "正常接通",
  "identity": {
    "name": null,
    "id_last4": null,
    "self_claim": "未明确"
  },
  "vote_intent": {
    "option1": null,
    "option2": null,
    "full_text": null
  },
  "explanation": "通话接通，但业主仅表示在忙稍后回电，未进行任何投票交流。"
}

【要求】
- 只输出JSON，不要任何解释、前缀、后缀或Markdown代码块标记（如```json）。
- 如果信息缺失，对应字段设为null。
- 确保JSON格式合法，字段名和字符串值使用双引号。
"""

def read_excel_vote_data(excel_path):
    """读取Excel，返回字典：{票号: [业主信息1, 业主信息2, ...]}"""
    df = pd.read_excel(excel_path)
    vote_dict = {}
    for _, row in df.iterrows():
        vote_no = str(row[EXCEL_KEY_COLUMNS["票号"]]).strip()
        owner_info = {
            "姓名": str(row[EXCEL_KEY_COLUMNS["业主姓名"]]).strip(),
            "身份证": str(row[EXCEL_KEY_COLUMNS["身份证号"]]).strip(),
            "投票原文": str(row[EXCEL_KEY_COLUMNS["投票内容"]]).strip() if pd.notna(row[EXCEL_KEY_COLUMNS["投票内容"]]) else ""
        }
        if vote_no not in vote_dict:
            vote_dict[vote_no] = []
        vote_dict[vote_no].append(owner_info)
    return vote_dict

def read_all_txt_files(folder_path):
    """读取文件夹内所有txt文件，返回按票号分组的字典：{票号: [文件信息列表]}"""
    files_by_vote = {}
    for filename in os.listdir(folder_path):
        if filename.endswith(".txt"):
            # 提取票号（第一个下划线前的内容）
            vote_no = filename.split("_")[0]
            # 提取时间戳（第二个部分，格式如20260118091256）
            parts = filename.split('_')
            timestamp_str = parts[1] if len(parts) > 1 else "0"
            try:
                timestamp = datetime.strptime(timestamp_str, "%Y%m%d%H%M%S")
            except:
                timestamp = datetime.min  # 解析失败则设为最小值
            file_path = os.path.join(folder_path, filename)
            with open(file_path, encoding="utf-8") as f:
                content = f.read().strip()
            file_info = {
                "文件名": filename,
                "时间戳": timestamp,
                "文本内容": content
            }
            if vote_no not in files_by_vote:
                files_by_vote[vote_no] = []
            files_by_vote[vote_no].append(file_info)
    return files_by_vote

def call_model_with_prompt(user_content):
    """使用新版MLX-LM API调用模型，返回原始响应字符串"""
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content}
    ]
    prompt = tokenizer.apply_chat_template(messages, tokenize=False, add_generation_prompt=True)
    
    # 创建采样器，temp=0 保证确定性输出
    sampler = make_sampler(temp=0)
    
    response = generate(
        model,
        tokenizer,
        prompt=prompt,
        max_tokens=1200,          # 增加到1000，确保长文本完整输出
        sampler=sampler,
        verbose=False
    )
    return response

def clean_response(response):
    """清理响应中的控制字符和多余空格"""
    # 移除控制字符
    response = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', response)
    # 将多个空格替换为单个空格（保留换行以便JSON解析）
    response = re.sub(r'[ \t]+', ' ', response).strip()
    return response

def extract_json_from_response(response):
    """
    从模型响应中提取JSON对象，支持多种格式。
    返回解析后的字典，或None
    """
    # 先清理响应
    response = clean_response(response)
    
    # 尝试直接解析整个响应
    try:
        return json.loads(response)
    except:
        pass
    
    # 尝试匹配被 ```json``` 包裹的代码块
    code_block_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', response, re.DOTALL | re.IGNORECASE)
    if code_block_match:
        try:
            return json.loads(code_block_match.group(1))
        except:
            pass
    
    # 尝试查找第一个 { 到最后一个 } 之间的内容（考虑嵌套括号）
    stack = []
    start = -1
    end = -1
    for i, ch in enumerate(response):
        if ch == '{':
            if not stack:
                start = i
            stack.append(ch)
        elif ch == '}':
            if stack:
                stack.pop()
                if not stack:
                    end = i
                    json_str = response[start:end+1]
                    try:
                        return json.loads(json_str)
                    except:
                        # 如果解析失败，重置并继续寻找下一个可能的JSON
                        start = -1
                        end = -1
    return None

def parse_model_response(response, excel_owners, vote_no=None, filename=None):
    """
    解析模型返回的JSON，并尝试匹配业主。
    返回：解析后的字典（包含matched_owner字段和best_score）或None（解析失败）
    """
    # 提取JSON
    res = extract_json_from_response(response)
    if res is None:
        # 确保日志目录存在
        log_dir = os.path.dirname(LOG_FILE_PATH)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        # 调试：保存原始响应到日志（带上票号和文件名）
        with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
            f.write(f"票号: {vote_no} 文件: {filename} 解析失败的响应:\n{response}\n{'-'*50}\n")
        return None

    # 确保关键字段存在且为字典
    if "call_status" not in res:
        res["call_status"] = "未知"
    
    # 处理 identity
    identity = res.get("identity")
    if identity is None or not isinstance(identity, dict):
        identity = {}
        res["identity"] = identity
    # 确保 identity 子字段存在
    identity.setdefault("name", None)
    identity.setdefault("id_last4", None)
    identity.setdefault("self_claim", "未明确")
    
    # 处理 vote_intent
    vote_intent = res.get("vote_intent")
    if vote_intent is None or not isinstance(vote_intent, dict):
        vote_intent = {}
        res["vote_intent"] = vote_intent
    vote_intent.setdefault("option1", None)
    vote_intent.setdefault("option2", None)
    vote_intent.setdefault("full_text", None)
    
    # 身份匹配
    matched_owner = None
    extracted_name = identity.get("name")
    extracted_id = identity.get("id_last4")
    best_score = 0   # 记录匹配最高分

    if excel_owners and (extracted_name or extracted_id):
        best_score = 0
        best_owner = None
        for owner in excel_owners:
            score = 0
            # 姓名相似度（拼音 + 原始文本）
            if extracted_name and owner["姓名"]:
                # 拼音相似度
                owner_py = ''.join(lazy_pinyin(owner["姓名"]))
                ext_py = ''.join(lazy_pinyin(extracted_name))
                score += fuzz.ratio(owner_py, ext_py) * 0.6
                # 原始文本编辑距离
                score += fuzz.ratio(owner["姓名"], extracted_name) * 0.4
            # 身份证后四位匹配（优先级高）
            if extracted_id and owner["身份证"].endswith(extracted_id):
                score += 100   # 直接拉高分数
            if score > best_score:
                best_score = score
                best_owner = owner
        # 设定阈值，60分以上认为匹配成功
        if best_score >= 60:
            matched_owner = best_owner

    res["matched_owner"] = matched_owner
    res["best_score"] = best_score   # 保存匹配分数，用于评分
    return res

def score_file_result(parsed):
    """
    对单个文件的解析结果进行评分，用于排序选择最佳文件。
    返回评分元组 (priority, timestamp)
    优先级数值越小越优先。
    """
    if parsed is None:
        return (5, datetime.min)  # 解析失败优先级最低（5）
    
    status = parsed.get("call_status")
    if status != "正常接通":
        # 异常通话（未接通、忙音等）优先级4
        return (4, parsed.get("_timestamp", datetime.min))
    
    intent = parsed.get("vote_intent", {})
    if not isinstance(intent, dict):
        intent = {}
    has_intent = 1 if (intent.get("option1") or intent.get("option2")) else 0
    
    has_identity = 1 if parsed.get("matched_owner") is not None else 0
    
    if has_intent and has_identity:
        priority = 1  # 有效投票
    elif has_intent:
        priority = 2  # 有投票意愿但身份未匹配
    else:
        priority = 3  # 正常接通但无投票意愿
    return (priority, parsed.get("_timestamp", datetime.min))

def is_valid_vote(parsed):
    """
    根据解析结果判断投票是否有效。
    返回：(is_valid, reason)
    """
    if parsed is None:
        return False, "模型输出解析失败"
    
    # 通话状态异常直接无效
    if parsed.get("call_status") != "正常接通":
        return False, f"通话状态：{parsed.get('call_status')}"
    
    # 必须有投票意愿（option1或option2非空）
    intent = parsed.get("vote_intent", {})
    if not isinstance(intent, dict):
        intent = {}
    if not intent.get("option1") and not intent.get("option2"):
        return False, "无明确投票意愿"
    
    # 必须匹配到具体业主
    if parsed.get("matched_owner") is None:
        return False, "无法匹配业主身份"
    
    # 可选：与人工备注进行一致性检查（如果备注存在）
    owner = parsed["matched_owner"]
    manual_note = owner.get("投票原文", "")
    if manual_note:
        # 简单检查提取的关键词是否出现在备注中
        keywords = []
        if intent.get("option1"):
            keywords.append(intent["option1"])
        if intent.get("option2"):
            keywords.append(intent["option2"])
        match_found = any(kw in manual_note for kw in keywords if kw)
        if not match_found:
            # 不一致，但仍算有效，但标记风险
            return True, "有效但人工备注不一致"
    
    return True, "有效投票"

def compute_vote_score(parsed):
    """
    根据解析结果计算投票评分（0~100分），用于人工复核。
    评分维度：
      1. 基础分（基于优先级）        0~60分
      2. 身份匹配质量分（best_score） 0~25分
      3. 选项明确性分                0~15分
      4. 人工备注一致性调整          -10~0分
    """
    if parsed is None:
        return 0.0

    # 1. 获取优先级并映射基础分
    priority, _ = score_file_result(parsed)
    if priority == 1:
        base = 60
    elif priority == 2:
        base = 45
    elif priority == 3:
        base = 30
    elif priority == 4:
        base = 15
    else:
        base = 0

    # 2. 身份匹配质量分（上限25分）
    best_score = parsed.get("best_score", 0)
    identity_score = min(best_score / 4, 25.0)

    # 3. 选项明确性分
    intent = parsed.get("vote_intent", {})
    has_option1 = intent.get("option1") is not None and intent.get("option1") != ""
    has_option2 = intent.get("option2") is not None and intent.get("option2") != ""
    if has_option1 and has_option2:
        clarity = 15
    elif has_option1 or has_option2:
        clarity = 8
    else:
        clarity = 0

    # 4. 人工备注一致性调整
    consistency_penalty = 0
    matched_owner = parsed.get("matched_owner")
    if matched_owner:
        manual_note = matched_owner.get("投票原文", "")
        if manual_note:
            # 检查模型提取的关键词是否出现在备注中
            keywords = []
            if intent.get("option1"):
                keywords.append(intent["option1"])
            if intent.get("option2"):
                keywords.append(intent["option2"])
            if keywords:
                match_found = any(kw in manual_note for kw in keywords if kw)
                if not match_found:
                    consistency_penalty = -10   # 完全不匹配，扣10分
            else:
                # 模型无提取，但备注非空，也可能有信息，扣5分
                consistency_penalty = -5

    total = base + identity_score + clarity + consistency_penalty
    return max(0.0, min(100.0, total))

def process_vote_no(vote_no, file_list, excel_owners):
    """
    处理单个票号下的所有文件，选择最佳结果。
    返回最终结果字典。
    """
    file_results = []
    for file_info in file_list:
        filename = file_info["文件名"]
        txt_content = file_info["文本内容"]
        
        # 构建用户消息（最后强调直接输出JSON，禁止思考）
        user_content = f"录音文本：{txt_content}\n"
        owner_names = [o["姓名"] for o in excel_owners] if excel_owners else []
        user_content += f"可能的业主：{', '.join(owner_names)}\n" if owner_names else ""
        user_content += "不要输出任何思考过程，直接输出JSON，不要包含```json标记，也不要任何解释。"
        
        # 调用模型
        model_response = call_model_with_prompt(user_content)
        
        # 可选调试：打印部分响应（可注释掉以保持输出简洁）
        # print(f"--- 文件 {filename} 响应预览 ---")
        # print(model_response[:300] + "..." if len(model_response) > 300 else model_response)
        # print("-----------------------------")
        
        # 解析并匹配（传入 vote_no 和 filename 以便日志记录）
        parsed = parse_model_response(model_response, excel_owners, vote_no=vote_no, filename=filename)
        if parsed:
            parsed["_timestamp"] = file_info["时间戳"]  # 附加时间戳用于排序
        
        file_results.append({
            "文件名": filename,
            "时间戳": file_info["时间戳"],
            "parsed": parsed
        })
    
    # 根据优先级排序，选择最佳文件
    # 排序规则：先按priority升序，同priority按时间戳降序（最新优先）
    file_results.sort(key=lambda x: score_file_result(x["parsed"]))
    best = file_results[0] if file_results else None
    
    if best is None:
        return {
            "票权编号": vote_no,
            "最终选用文件": "",
            "匹配业主": "",
            "投票状态": "无效",
            "风险备注": "无可用录音文件",
            "推理说明": "",
            "投票评分": 0,
            "模型输出": ""
        }
    
    parsed = best["parsed"]
    is_valid, reason = is_valid_vote(parsed)
    
    if parsed is None:
        matched_name = ""
        explanation = ""
        vote_score = 0
        model_output_dump = ""
    else:
        matched_owner = parsed.get("matched_owner")
        matched_name = matched_owner.get("姓名", "") if matched_owner else ""
        explanation = parsed.get("explanation", "")
        vote_score = compute_vote_score(parsed)
        # 创建副本并移除 _timestamp 字段，避免 JSON 序列化错误
        parsed_copy = parsed.copy()
        parsed_copy.pop("_timestamp", None)
        model_output_dump = json.dumps(parsed_copy, ensure_ascii=False)
    
    return {
        "票权编号": vote_no,
        "最终选用文件": best["文件名"],
        "匹配业主": matched_name,
        "投票状态": "有效" if is_valid else "无效",
        "风险备注": reason,
        "推理说明": explanation,
        "投票评分": round(vote_score, 1),
        "模型输出": model_output_dump
    }

def main():
    print("===== 业主大会投票核验系统（终极版 - 含评分）=====\n")
    
    # 1. 读取Excel
    print("正在读取Excel数据...")
    vote_data = read_excel_vote_data(EXCEL_PATH)
    print(f"共读取 {len(vote_data)} 个票号（含多业主）")
    
    # 2. 读取所有txt文件并按票号分组
    files_by_vote = read_all_txt_files(TXT_FOLDER)
    print(f"待核验录音文件：{sum(len(v) for v in files_by_vote.values())} 个，共 {len(files_by_vote)} 个票号\n")
    
    final_results = []
    
    # 3. 遍历每个票号
    for vote_no, file_list in tqdm(files_by_vote.items(), desc="核验进度", colour="green"):
        tqdm.write(f"🔍 票权编号：{vote_no} | 共有 {len(file_list)} 个录音文件")
        
        # 获取该票号下的所有业主
        excel_owners = vote_data.get(vote_no, [])
        
        # 处理该票号下的所有文件，选择最佳结果
        result = process_vote_no(vote_no, file_list, excel_owners)
        final_results.append(result)
        
        tqdm.write(f"✅ 最终选用文件：{result['最终选用文件']}")
        tqdm.write(f"✅ 匹配业主：{result['匹配业主']}")
        tqdm.write(f"📊 投票状态：{result['投票状态']}")
        tqdm.write(f"🎯 投票评分：{result['投票评分']}")
        tqdm.write(f"⚠️ 风险备注：{result['风险备注']}\n")
    
    # 4. 导出Excel
    df = pd.DataFrame(final_results)
    output_columns = ["票权编号", "最终选用文件", "匹配业主", "投票状态", "投票评分", "风险备注", "推理说明", "模型输出"]
    df[output_columns].to_excel(RESULT_SAVE_PATH, index=False)
    print(f"===== ✅ 全部完成！结果已保存至：{RESULT_SAVE_PATH} =====")

if __name__ == "__main__":
    main()