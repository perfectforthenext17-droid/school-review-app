import streamlit as st
from docx import Document
import re
from datetime import datetime
from zhipuai import ZhipuAI
import base64
import zipfile
import json
import io
from PIL import Image


# ================= 配置区 =================
client = ZhipuAI(api_key="ac67cb8c61804fe6bdcecdc13e8aee0a.v7IK0Kxs4EMLsMbr")

# ================= 初始化系统记忆 (Session State) =================
# 让网页记住每个班级的状态（颜色）和提交次数
if 'class_data' not in st.session_state:
    st.session_state.class_data = {}
    # 定义每个年级的班级数量
    grade_config = {"高一": 38, "高二": 38, "高三": 37}
    
    for grade, count in grade_config.items():
        st.session_state.class_data[grade] = {}
        for i in range(1, count + 1):
            st.session_state.class_data[grade][f"{i}班"] = {
                "status": "white",  # white: 未提交, green: 通过, red: 不通过
                "attempts": 0       # 记录提交次数
            }

# 记住当前正在操作哪个班级
if 'current_selected_class' not in st.session_state:
    st.session_state.current_selected_class = None

# ================= 工具函数 =================
import io

def extract_images_from_docx(docx_file):
    """从 docx 文件中提取所有图片 (彻底修复抓取到空文件夹的 Bug)"""
    images = []
    # 使用 io.BytesIO 将上传的文件流彻底读入内存
    file_bytes = docx_file.read()
    
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as docx_zip:
            for item in docx_zip.namelist():
                # 🛡️ 新增两道安检：必须在这个路径下，且名字绝对不能以 "/" 结尾（排除文件夹本身）
                if item.startswith('word/media/') and not item.endswith('/'):
                    try:
                        img_data = docx_zip.read(item)
                        # 🛡️ 第三道安检：确保读出的数据不是 0 字节的空图
                        if len(img_data) > 0:
                            images.append(img_data)
                    except zipfile.BadZipFile:
                        continue # 遇到坏图直接跳过
    except Exception as e:
        raise Exception(f"无法解析文档压缩结构，文件可能已损坏: {e}")
        
    return images

def check_time_duration(time_str):
    """
    宽松匹配时间段 (规则3优化版)
    支持：中英文冒号(:/：)、中英文连接符(-/—/~/～/至/到)、一位或两位数小时(6:19 或 06:19)
    """
    # 正则提取形如: 时:分 连接符 时:分
    # (\d{1,2})[:：](\d{1,2}) 匹配 开始时:分
    # [^\d]+ 匹配中间任意非数字字符（支持 -, ~, ～, 到, 至 等任意符号及空格）
    # (\d{1,2})[:：](\d{1,2}) 匹配 结束时:分
    pattern = r'(\d{1,2})\s*[:：]\s*(\d{1,2})\s*[^\d]+\s*(\d{1,2})\s*[:：]\s*(\d{1,2})'
    match = re.search(pattern, time_str)
    
    if not match:
        return False, f"未识别到有效的时间段格式（填写内容为：{time_str}），请确保包含类似 12:26-12:54 或 6:19～6:36 的时间范围。"
    
    h1, m1, h2, m2 = map(int, match.groups())
    
    # 基础数值合法性校验
    if not (0 <= h1 <= 24 and 0 <= m1 < 60 and 0 <= h2 <= 24 and 0 <= m2 < 60):
        return False, f"时间数值超出常规范围（识别为 {h1}:{m1} 至 {h2}:{m2}），请核对填写数字。"
    
    # 转换为当日总分钟数计算差值
    start_minutes = h1 * 60 + m1
    end_minutes = h2 * 60 + m2
    
    # 计算差值（考虑可能刚好跨午夜的情况）
    diff = end_minutes - start_minutes
    if diff < 0:
        diff += 24 * 60  # 如果跨越了午夜（例如 23:50 至 00:10）
        
    if diff < 10:
        return False, f"活动时长不足 10 分钟（识别为 {h1}:{m1:02d} 至 {h2}:{m2:02d}，实际共 {diff} 分钟）。"
        
    return True, f"时间合格（时长共 {diff} 分钟）"

import base64
import json

def review_images_with_ai(images):
    """调用智谱 GLM-4V 检查图片 (四宫格拼图模式，彻底解决传图报错)"""
    print("\n" + "="*40)
    print("▶️ 启动 AI 视觉审查模块 (四宫格拼图模式)...")
    
    try:
        # 1. 把提取到的 4 张图片的字节数据变成图像对象
        pil_images = []
        for img_bytes in images:
            # 打开图片并强制转为标准 RGB 格式
            img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
            # 把每张图统一缩放成 600x600 的正方形，方便拼接
            img = img.resize((600, 600))
            pil_images.append(img)
        
        # 2. 创建一张 1200 x 1200 的纯白大背景图
        grid_img = Image.new('RGB', (1200, 1200), color='white')
        
        # 3. 把 4 张图片按坐标贴到四个角 (左上, 右上, 左下, 右下)
        positions = [(0, 0), (600, 0), (0, 600), (600, 600)]
        for i, img in enumerate(pil_images[:4]):
            grid_img.paste(img, positions[i])
            
        print("📦 4张图片已成功拼接为 1 张四宫格大图！")
        
        # 4. 把拼接好的大图转成智谱要求的【纯 Base64 格式】（不带前缀）
        buffer = io.BytesIO()
        grid_img.save(buffer, format="JPEG", quality=85)
        base64_img = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        # 5. 构建给 AI 的命令 (现在只发一张大图)
        content_list = [
            {
                "type": "text",
                "text": """
                你现在是一个严格的审核员。我给你发送了一张由4张活动照片拼接成的“四宫格大图”。
                排版顺序为：左上是图1，右上是图2，左下是图3，右下是图4。
                
                请执行以下检查：
                1. 水印检查：请仔细检查这4个分格画面，是否每个画面都带有包含时间（年、月、日、时、分）的水印？如果任何一个画面没有，则不通过。
                2. 人脸检查：右下角的图4是参加人员合照。请计算右下角图4中清晰可见的人脸数量，必须大于3个（至少4个人脸）。
                
                请务必仅返回如下 JSON 格式：
                {"watermark_passed": true/false, "faces_passed": true/false, "reason": "你的判断理由"}
                """
            },
            {
                "type": "image_url",
                "image_url": {
                    "url": base64_img  # 严格使用纯 base64，去掉了 data:image 前缀
                }
            }
        ]
        
        print("🚀 正在向智谱服务器发送拼接大图 (设置40秒防卡死超时保护)...")
        response = client.chat.completions.create(
            model="glm-4v",
            messages=[{"role": "user", "content": content_list}],
            temperature=0.1,
            timeout=40
        )
        
        print("✅ 成功收到 AI 服务器回传的数据！")
        result_text = response.choices[0].message.content
        result_text = result_text.replace("```json", "").replace("```", "").strip()
        print(f"🤖 AI 原始判定结果: {result_text}")
        
        result = json.loads(result_text)
        passed = result.get("watermark_passed", False) and result.get("faces_passed", False)
        reason = result.get("reason", "未知原因")
        
        print("🏁 AI 模块执行完毕！")
        print("="*40 + "\n")
        return passed, reason
        
    except Exception as e:
        error_msg = f"AI 视觉审查失败，错误详情: {e}"
        print(f"❌ 捕获到错误: {error_msg}")
        return False, error_msg
# ================= 网页前端 UI =================

st.title("📋 团队活动自动化审查系统")
st.header("志愿服务活动审查")

# 1. 选择年级
grade_choice = st.radio("请选择活动所属年级：", ["高一", "高二", "高三"])

# 当切换年级时，清空当前选中的班级，避免高一1班串到高二1班
if 'last_grade' not in st.session_state:
    st.session_state.last_grade = grade_choice
if st.session_state.last_grade != grade_choice:
    st.session_state.current_selected_class = None
    st.session_state.last_grade = grade_choice

st.divider()
st.subheader(f"📍 请选择{grade_choice}的盒子")

# 2. 渲染班级状态面板 (大方块勾选项)
classes_dict = st.session_state.class_data[grade_choice]

# 每行显示 6 个班级按钮，整齐排列
cols = st.columns(6)
color_map = {"white": "⬜", "green": "🟩", "red": "🟥"}

for idx, (cls_name, data) in enumerate(classes_dict.items()):
    col = cols[idx % 6] # 循环放在6列中
    icon = color_map[data["status"]]
    attempts = data["attempts"]
    
    # 按钮上显示状态方块、班级名和剩余次数
    button_label = f"{icon} {cls_name} ({attempts}/3)"
    
    # 用户点击某个班级按钮
    if col.button(button_label, key=f"btn_{grade_choice}_{cls_name}", use_container_width=True):
        st.session_state.current_selected_class = cls_name

st.divider()

# 3. 提交文件环节 (仅当选中班级后才显示)
if st.session_state.current_selected_class:
    current_class = st.session_state.current_selected_class
    class_info = st.session_state.class_data[grade_choice][current_class]
    
    st.subheader(f"📤 当前操作：{grade_choice} {current_class}")
    
    # 检查是否还能提交
    if class_info["status"] == "green":
        st.success("✅ 该班级文件已审查通过，无需重复提交！")
    elif class_info["attempts"] >= 3:
        st.error("🚫 该班级已达到最大提交次数（3次），审查通道已锁定。")
    else:
        st.info(f"剩余提交次数：{3 - class_info['attempts']} 次")
        uploaded_file = st.file_uploader(f"请上传 {current_class} 的《志愿服务活动模板.docx》", type="docx", key="uploader")

        if uploaded_file is not None:
            st.warning("正在执行自动化审查，请稍候...")
            errors = []
            
            try:
                # --- 开始执行你之前的审查规则 ---
                # --- 开始执行自动化审查规则 ---
                doc = Document(uploaded_file)
                participants_count = 0
                time_duration_str = ""
                
                # 遍历文档所有的段落
                for para in doc.paragraphs:
                    # 1. 获取最原始的文字
                    raw_text = para.text
                    
                    # 2. 【文本净化器】：无视排版、无视缩进、无视字体
                    # 把所有的空格、缩进(Tab)、换行全部删掉
                    clean_text = re.sub(r'\s+', '', raw_text)
                    # 将全角的冒号、顿号统一替换或删除，方便直接搜索汉字
                    clean_text = clean_text.replace("：", ":").replace("、", "")
                    
                    # 3. 极度宽松的模糊匹配
                    # 只要净化后的文本里包含“参与人数”这四个字，不管前面是“二”还是“2”，全算数
                    if "参与人数" in clean_text:
                        # 提取这四个字后面的第一组数字
                        nums = re.findall(r'参与人数.*?(\d+)', clean_text)
                        if nums: 
                            participants_count = int(nums[0])
                            
                    # 只要净化后的文本包含“活动时长”这四个字，就把原始文本交给咱们之前写好的极度包容的时间判定函数
                    elif "活动时长" in clean_text:
                        time_duration_str = raw_text
                        
                # 规则 1 & 2
                if participants_count == 0:
                     errors.append("未读取到参与人数。")
                else:
                    if grade_choice == "高一" and not (10 <= participants_count <= 12):
                        errors.append(f"高一人数应为10-12人，实际 {participants_count} 人。")
                    elif grade_choice in ["高二", "高三"] and participants_count < 10:
                        errors.append(f"{grade_choice}人数应≥10人，实际 {participants_count} 人。")
                        
                # 规则 3
                if not time_duration_str:
                    errors.append("未读取到活动时长。")
                else:
                    time_passed, time_msg = check_time_duration(time_duration_str)
                    if not time_passed: errors.append(time_msg)
                        
                # 规则 4
                if not doc.tables:
                    errors.append("缺失人员表格。")
                else:
                    table = doc.tables[0]
                    valid_person_count = 0
                    for row in table.rows[1:]:
                        cells_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                        phones = [text for text in cells_text if re.fullmatch(r'\d{11}', text)]
                        valid_person_count += len(phones)
                    if valid_person_count != participants_count:
                        errors.append(f"表格手机号数量({valid_person_count})与参与人数({participants_count})不符。")

                # 规则 5 (图片数量预检)
                # 规则 5 (图片数量预检与 AI 深度审查)
                uploaded_file.seek(0)
                images = extract_images_from_docx(uploaded_file)
                if len(images) != 4:
                     errors.append(f"图片数量不正确，必须为 4 张，实际提取到 {len(images)} 张。")
                else:
                     # 数量没问题，呼叫 AI 上场！
                     st.info("✅ 成功提取 4 张图片，正在呼叫 AI 进行深度视觉审查，请稍候约10秒...")
                     ai_passed, ai_reason = review_images_with_ai(images)
                     
                     if not ai_passed:
                         errors.append(f"AI 图片审查未通过：{ai_reason}")
                     else:
                         st.success(f"🤖 AI 审查判定：{ai_reason}")

                # --- 审查结果判定与状态更新 ---
                class_info["attempts"] += 1 # 扣除一次提交次数
                
                if len(errors) == 0:
                    class_info["status"] = "green" # 标记为通过
                    st.success(f"🎉 {current_class} 审查通过！")
                else:
                    class_info["status"] = "red" # 标记为不通过
                    st.error("❌ 审查未通过！原因如下：")
                    for i, err in enumerate(errors):
                        st.write(f"{i+1}. {err}")
                
                # 强制网页刷新，让上方的班级看板立刻更新颜色
                if st.button("点击刷新看板颜色"):
                    st.rerun()

            except Exception as e:
                st.error(f"解析文件出错: {e}")