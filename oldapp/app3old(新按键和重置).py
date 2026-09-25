import os
os.environ["HTTP_PROXY"] = ""
os.environ["HTTPS_PROXY"] = ""

import streamlit as st
import re
import streamlit as st
import re
from docx import Document
from datetime import datetime
from zhipuai import ZhipuAI
import base64
import json
import io
from PIL import Image
import zipfile
import time
from supabase import create_client
import gc

# ================= 全局屏蔽 ZIP CRC-32 严格校验 =================
# 彻底解决手机端上传 Word 文档时，由于文件流微小重组导致的 Bad CRC-32 崩溃问题
_original_init = zipfile.ZipExtFile.__init__
def _patched_init(self, *args, **kwargs):
    _original_init(self, *args, **kwargs)
    self._expected_crc = None  # 强行抹除预期校验码，强制放行所有微损文件！
zipfile.ZipExtFile.__init__ = _patched_init
# ========================================================================
# ================= 配置区 =================
client = ZhipuAI(api_key=st.secrets["ZHIPU_API_KEY"])

# 🆕 Supabase 云端配置 (替换为你的真实凭证)
# 🆕 Supabase 云端配置 (安全读取 secrets)
SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
sb = create_client(SUPABASE_URL, SUPABASE_KEY)

# ================= 1. 初始化系统记忆 =================
if 'admin_logged_in' not in st.session_state:
    st.session_state.admin_logged_in = False
if 'grade_config' not in st.session_state:
    st.session_state.grade_config = {"高一": 38, "高二": 38, "高三": 37}
if 'current_selected_class' not in st.session_state:
    st.session_state.current_selected_class = None

# ================= 2. 工具函数与状态查询 =================
# 动态获取当前正在进行的批次
# 🆕 使用 st.cache_data 缓存查询结果，将 1-2秒 的网络延迟压缩到 0.01秒
@st.cache_data(ttl=60, show_spinner=False)
def get_active_batch():
    batch_res = sb.table("batches").select("batch_name").eq("is_active", 1).order("id", desc=True).limit(1).execute()
    return batch_res.data[0]['batch_name'] if batch_res.data else None

st.session_state.current_batch = get_active_batch()

@st.cache_data(ttl=15, show_spinner=False)
def get_class_status_from_db(batch, activity_type, grade):
    if not batch: return {}
    res = sb.table("submissions").select("class_name, status, attempts").match({
        "batch_name": batch,
        "activity_type": activity_type,
        "grade": grade
    }).execute()
    return {r['class_name']: {"status": r['status'], "attempts": r['attempts']} for r in res.data}

def extract_images_from_docx(docx_file):
    images = []
    file_bytes = docx_file.read()
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as docx_zip:
            for item in docx_zip.namelist():
                # 兼容性升级：仅提取标准图像格式，强力拦截桌面端 Word 生成的 .emf/.wmf 矢量图或 .bin 缓存文件
                if item.startswith('word/media/') and item.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.webp')):
                    try:
                        img_data = docx_zip.read(item)
                        if len(img_data) > 0: images.append(img_data)
                    except Exception: 
                        # 捕捉所有底层解压报错，兼容手机端微损图片
                        continue
    except Exception as e:
        raise Exception(f"文档结构损坏: {e}")
    return images
def check_time_duration(time_str):
    pattern = r'(\d{1,2})\s*[:：]\s*(\d{1,2})\s*[^\d]+\s*(\d{1,2})\s*[:：]\s*(\d{1,2})'
    match = re.search(pattern, time_str)
    if not match: return False, "未识别到有效时间段格式(需类似 15:36~15:48)"
    
    h1, m1, h2, m2 = map(int, match.groups())
    start = h1 * 60 + m1
    end = h2 * 60 + m2
    
    # 修复：先老老实实计算差值，再做判断
    diff = end - start
    if diff < 0: 
        diff += 24 * 60
        
    if diff < 10: 
        return False, f"时长不足10分钟(实际{diff}分钟)"
        
    return True, "时间合格"

def review_images_with_ai(images):
    """调用智谱 AI (glm-4v) 进行四宫格图片审查（加入极限内存释放机制）"""
    try:
        pil_images = []
        for img_bytes in images:
            # 使用 with 语句确保原始图片流在使用后立即关闭
            with Image.open(io.BytesIO(img_bytes)) as img:
                pil_images.append(img.convert("RGB").resize((600, 600)))
        
        grid_img = Image.new('RGB', (1200, 1200), color='white')
        positions = [(0, 0), (600, 0), (0, 600), (600, 600)]
        
        for i, img in enumerate(pil_images[:4]): 
            grid_img.paste(img, positions[i])
            # 贴完图后立刻关闭单个图片对象释放内存
            img.close()
            
        buffer = io.BytesIO()
        grid_img.save(buffer, format="JPEG", quality=85)
        base64_img = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        # ⚠️ 核心操作：大图片转码完成后，立刻物理销毁画布和缓存对象
        grid_img.close()
        buffer.close()
        del pil_images
        del grid_img
        gc.collect() # 强制呼叫系统回收垃圾内存
        
        content_list = [
            {"type": "text", "text": """
            审核4张照片拼成的四宫格。
            1. 每一张必须有时间水印(xx年xx日xx点xx分)。
            2. 右下角的图4必须是全员合照，人脸数必须大于3。
            返回JSON: {"watermark_passed": true/false, "faces_passed": true/false, "reason": "理由"}
            """},
            {"type": "image_url", "image_url": {"url": base64_img}}
        ]
        
        response = client.chat.completions.create(
            model="glm-4v", messages=[{"role": "user", "content": content_list}], temperature=0.1, timeout=40
        )
        
        # 释放 Base64 长字符串
        del base64_img
        gc.collect()
        
        result_text = response.choices[0].message.content.replace("```json", "").replace("```", "").strip()
        result = json.loads(result_text)
        return result.get("watermark_passed", False) and result.get("faces_passed", False), result.get("reason", "未知")
    except Exception as e:
        return False, f"AI审查失败: {e}"
# ================= 3. 侧边栏导航 =================
with st.sidebar:
    st.title("系统导航")
    page_mode = st.radio("请选择访问入口：", ["🟢 学生提交端", "⚙️ 管理员后台"])
    st.divider()
    st.caption(f"当前活动批次：{st.session_state.current_batch}")

# ================= 4. ⚙️ 管理员后台逻辑 =================
if page_mode == "⚙️ 管理员后台":
    st.title("⚙️ 团队活动中央控制台")
    if not st.session_state.admin_logged_in:
        pwd = st.text_input("管理员密码", type="password")
        if st.button("登录"):
            if pwd == "qiushihao2009":
                st.session_state.admin_logged_in = True
                st.rerun()
            else: st.error("密码错误！")
    else:
        if st.sidebar.button("登出系统"):
            st.session_state.admin_logged_in = False
            st.rerun()
            
        t1, t2, t3, t4 = st.tabs(["📊 综合统计", "📸 团日审核", "📦 导出", "⚙️ 设置"])
        
        with t1: 
            st.subheader(f"📊 {st.session_state.current_batch} - 综合统计大屏")
            stat_activity = st.radio("请选择要查看的活动类型：", ["🚩 团日活动", "🎈 志愿服务活动"], horizontal=True)
            current_stat_type = "团日活动" if "团日活动" in stat_activity else "志愿服务活动"
            
            # 🆕 替换为 Supabase 查询
            res = sb.table("submissions").select("grade, class_name, status").match({
                "batch_name": st.session_state.current_batch,
                "activity_type": current_stat_type
            }).execute()
            rows = res.data
            
            db_status_map = {f"{r['grade']}_{r['class_name']}": r['status'] for r in rows}
            st.divider()
            
            for grade in ["高一", "高二", "高三"]:
                total_classes = st.session_state.grade_config[grade]
                passed, pending, failed, unsubmitted = [], [], [], []
                
                for i in range(1, total_classes + 1):
                    cls_name = f"{i}班"
                    status = db_status_map.get(f"{grade}_{cls_name}", "white") 
                    if status == "green": passed.append(cls_name)
                    elif status == "pending": pending.append(cls_name)
                    elif status == "red": failed.append(cls_name)
                    else: unsubmitted.append(cls_name)
                    
                st.markdown(f"#### 📍 {grade} (已交: {len(passed) + len(pending)} / {total_classes} 班)")
                c1, c2, c3, c4 = st.columns(4)
                with c1: st.success(f"✅ 合格 ({len(passed)})\n\n" + "、".join(passed))
                with c2: st.warning(f"⏳ 待审 ({len(pending)})\n\n" + "、".join(pending))
                with c3: st.error(f"❌ 被驳回 ({len(failed)})\n\n" + "、".join(failed))
                with c4: 
                    st.info(f"⬜ 未交 ({len(unsubmitted)})")
                    with st.expander("名单"): st.write("、".join(unsubmitted))
                st.divider()

        with t2: 
            st.subheader("📸 团日活动 - 人工视觉审核池")
            # 🆕 替换为 Supabase 查询待审核列表
            pending_res = sb.table("submissions").select("*").match({
                "activity_type": "团日活动",
                "status": "pending"
            }).order("created_at").execute()
            pending_rows = pending_res.data
            
            if len(pending_rows) == 0:
                st.success("🎉 太棒了！当前没有任何待审核的团日活动。")
            else:
                st.info(f"当前共有 {len(pending_rows)} 个班级等待人工审核。")
                if st.button("⚡ 一键通过所有待审班级 (特权操作)", type="primary"):
                    sb.table("submissions").update({"status": "green"}).match({"activity_type": "团日活动", "status": "pending"}).execute()
                    st.success("已清空所有待审任务，状态全部变更为合格！")
                    st.rerun()
                
                st.divider()
                current_task = pending_rows[0]
                task_id = current_task['id']
                st.write(f"正在审核：**{current_task['batch_name']} - {current_task['grade']} {current_task['class_name']}**")
                
                images = [base64.b64decode(img) for img in current_task['images_data']]
                c1, c2, c3 = st.columns(3)
                if len(images) == 3:
                    c1.image(images[0], use_container_width=True)
                    c2.image(images[1], use_container_width=True)
                    c3.image(images[2], use_container_width=True)
                
                col_btn1, col_btn2 = st.columns([1, 1])
                if col_btn1.button("✅ 画面合规 (通过)", key="pass_btn", use_container_width=True):
                    sb.table("submissions").update({"status": "green"}).eq("id", task_id).execute()
                    get_class_status_from_db.clear() # 🆕 清空缓存
                    st.rerun()
                if col_btn2.button("❌ 画面违规 (驳回)", key="reject_btn", use_container_width=True):
                    sb.table("submissions").update({"status": "red"}).eq("id", task_id).execute()
                    get_class_status_from_db.clear() # 🆕 清空缓存
                    st.rerun()

        with t3: 
            st.subheader("📦 精准归档与分类导出")
            st.info("📱 **移动端管理员须知**：由于微信/QQ内置浏览器限制，若点击下方下载压缩包无反应，请点击右上角「···」选择“在系统浏览器打开”后重试。")
            
            res_batches = sb.table("batches").select("batch_name").order("id", desc=True).execute()
            batch_rows = res_batches.data
            
            if not batch_rows:
                st.warning("⚠️ 系统中尚未创建任何批次记录。")
            else:
                batch_list = [row['batch_name'] for row in batch_rows]
                
                col_batch, col_act, col_grade = st.columns(3)
                with col_batch:
                    export_batch = st.selectbox("📅 选择批次", batch_list)
                with col_act: 
                    export_activity = st.selectbox("📌 选择活动类型", ["🚩 团日活动", "🎈 志愿服务活动"])
                with col_grade: 
                    export_grade = st.selectbox("🎓 选择年级", ["高一", "高二", "高三"])
                    
                clean_activity = "团日活动" if "团日" in export_activity else "志愿服务活动"
                st.divider()
                
                # 动态生成当前任务的唯一标识，用于缓存状态隔离
                task_key = f"{export_batch}_{export_grade}_{clean_activity}"
                
                if st.button(f"⚙️ 立即打包：【{export_batch} - {export_grade} - {clean_activity}】", use_container_width=True, type="primary"):
                    
                    res_files = sb.table("submissions").select("class_name, file_path, original_filename").match({
                        "batch_name": export_batch,
                        "status": "green",
                        "activity_type": clean_activity,
                        "grade": export_grade
                    }).execute()
                    rows = res_files.data
                    
                    if len(rows) == 0:
                        st.warning(f"⚠️ 在【{export_batch}】批次下，没有找到【{export_grade} - {clean_activity}】的合格文件。")
                        if 'zip_data' in st.session_state: del st.session_state.zip_data
                    else:
                        with st.spinner("正在从云端飞速拉取文件并打包 (可能需要十几秒)..."):
                            zip_buffer = io.BytesIO()
                            failed_downloads = [] # 记录下载失败的班级
                            
                            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                                def get_class_num(row):
                                    num_str = row['class_name'].replace('班', '')
                                    return int(num_str) if num_str.isdigit() else 999
                                    
                                sorted_rows = sorted(rows, key=get_class_num)
                                
                                for r in sorted_rows:
                                    if r.get('file_path'):  
                                        try:
                                            # 容错：单独拉取文件，失败不影响全局
                                            doc_bytes = sb.storage.from_("school-docs").download(r['file_path'])
                                            
                                            # 防覆盖：强制在文件名前追加班级名称
                                            og_name = r.get('original_filename')
                                            if og_name:
                                                file_name = og_name if r['class_name'] in og_name else f"{r['class_name']}_{og_name}"
                                            else:
                                                file_name = f"{r['class_name']}.docx"
                                                
                                            zip_file.writestr(file_name, doc_bytes)
                                        except Exception as e:
                                            failed_downloads.append(r['class_name'])
                                            continue
                            
                            # 将生成的二进制流存入内存，彻底解决按钮嵌套断连问题
                            st.session_state.zip_data = zip_buffer.getvalue()
                            st.session_state.zip_name = f"{task_key}.zip"
                            st.session_state.zip_count = len(sorted_rows) - len(failed_downloads)
                            
                            if failed_downloads:
                                st.error(f"❌ 以下班级云端原文件拉取失败，未打包：{', '.join(failed_downloads)}")
                
                # 无论是否点击“打包”，只要内存里有数据就持续渲染下载按钮
                if st.session_state.get('zip_data') and st.session_state.get('zip_name') == f"{task_key}.zip":
                    st.success(f"🎉 打包成功！共提取了 {st.session_state.zip_count} 份文件，已按班级顺序排列。")
                    st.download_button(
                        label="⬇️ 点击下载压缩包", 
                        data=st.session_state.zip_data, 
                        file_name=st.session_state.zip_name, 
                        mime="application/zip", 
                        use_container_width=True
                    )
        with t4: 
            st.subheader("⚙️ 系统发布与历史归档")
            st.markdown("### 📢 当前收集项目")
            
            if st.session_state.current_batch:
                st.success(f"🟢 当前正在收集中：**{st.session_state.current_batch}**")
                
                if st.button("🛑 终止当前项目 (进入归档)", type="primary"):
                    # 🆕 替换为 Supabase 更新
                    sb.table("batches").update({"is_active": 0}).eq("batch_name", st.session_state.current_batch).execute()
                    st.session_state.current_batch = None
                    time.sleep(0.2) 
                    st.rerun()
            else:
                st.warning("⚪ 当前没有正在收集的项目。")
                with st.form("new_batch_form"):
                    new_batch = st.text_input("新建收集项目名称 (例如：2026年10月团日与志愿收集)：")
                    if st.form_submit_button("🚀 立即发布新项目"):
                        if new_batch:
                            try:
                                # 🆕 替换为 Supabase 插入
                                sb.table("batches").insert({"batch_name": new_batch, "is_active": 1}).execute()
                                st.session_state.current_batch = new_batch
                                time.sleep(0.2)
                                st.rerun()
                            except Exception:
                                st.error("该名称已存在，请更换一个项目名！")

            st.divider()
            st.markdown("### 🗄️ 历史项目记录")
            # 🆕 替换为 Supabase 读取历史批次
            hist_res = sb.table("batches").select("batch_name, created_at").eq("is_active", 0).order("id", desc=True).execute()
            history = hist_res.data
            
            if history:
                for h in history:
                    old_batch = h['batch_name']
                    with st.expander(f"📁 {old_batch} (发布于: {h['created_at']})"):
                        hist_activity = st.radio(f"查看 {old_batch} 的：", ["🚩 团日活动", "🎈 志愿服务活动"], key=f"radio_{old_batch}", horizontal=True)
                        clean_hist_act = "团日活动" if "团日" in hist_activity else "志愿服务活动"
                        
                        # 🆕 替换为 Supabase 读取历史详细状态
                        res_hist = sb.table("submissions").select("grade, class_name, status").match({
                            "batch_name": old_batch,
                            "activity_type": clean_hist_act
                        }).execute()
                        rows = res_hist.data
                        
                        hist_status_map = {f"{r['grade']}_{r['class_name']}": r['status'] for r in rows}
                        
                        for grade in ["高一", "高二", "高三"]:
                            total_classes = st.session_state.grade_config[grade]
                            passed, pending, failed, unsubmitted = [], [], [], []
                            
                            for i in range(1, total_classes + 1):
                                cls_name = f"{i}班"
                                status = hist_status_map.get(f"{grade}_{cls_name}", "white") 
                                if status == "green": passed.append(cls_name)
                                elif status == "pending": pending.append(cls_name)
                                elif status == "red": failed.append(cls_name)
                                else: unsubmitted.append(cls_name)
                                
                            st.markdown(f"**📍 {grade}** (已交: {len(passed) + len(pending)} / {total_classes} 班)")
                            c1, c2, c3, c4 = st.columns(4)
                            with c1: st.success(f"✅ 合格 ({len(passed)})")
                            with c2: st.warning(f"⏳ 待审 ({len(pending)})")
                            with c3: st.error(f"❌ 驳回 ({len(failed)})")
                            with c4: st.info(f"⬜ 未交 ({len(unsubmitted)})")
                            st.divider()
            else:
                st.caption("暂无历史记录。")
# ================= 5. 🟢 学生提交端逻辑 =================
elif page_mode == "🟢 学生提交端":
    
    # 🌟 1. 零延迟加载机制：在后端将校徽强转为 Base64 嵌入前端
    # 🌟 1. 零延迟加载机制：在后端将校徽强转为 Base64 嵌入前端
    import os, base64
    logo_base64 = ""
    logo_path = "dff2f6bd9341d59fef8359f9cf1556f7.jpg"
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            logo_base64 = base64.b64encode(f.read()).decode('utf-8')
            
    # 提前生成图片 HTML，彻底避开 VS Code 的嵌套 f-string 解析警告
    logo_html = f'<img src="data:image/jpeg;base64,{logo_base64}" class="school-logo">' if logo_base64 else ''

    # 🌟 2. 注入“纯血 iOS 极简风”全新 CSS (完美居中 + 高级质感滑块 + 校徽防丢失版)
    st.markdown(f"""
    <style>
    /* 🎯 全局系统字体 */
    html, body, [class*="st-"], h1, h2, h3, h4, h5, h6, p, label, button, div {{
        font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Helvetica Neue", Arial, sans-serif !important;
    }}

    :root {{
        --dj-red: #8B1C31; 
        --dj-red-light: rgba(139, 28, 49, 0.06); 
        --dj-red-shadow: rgba(139, 28, 49, 0.15);
        --dj-card-bg: var(--secondary-background-color); 
        --dj-text: var(--text-color);
        --dj-bg: var(--background-color);
    }}
    
    @media (prefers-color-scheme: dark) {{
        :root {{
            --dj-red: #ff768b; 
            --dj-red-light: rgba(255, 118, 139, 0.15);
            --dj-red-shadow: rgba(255, 118, 139, 0.25);
        }}
    }}

    @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(12px); }} to {{ opacity: 1; transform: translateY(0); }} }}
    .block-container {{ animation: fadeIn 0.6s cubic-bezier(0.16, 1, 0.3, 1); }}
    @keyframes elegantEntrance {{ from {{ opacity: 0; transform: scale(0.95) translateY(10px); }} to {{ opacity: 1; transform: scale(1) translateY(0); }} }}
    
    @keyframes textPop {{
        0% {{ opacity: 0; transform: translateY(10px) scale(0.9); color: var(--dj-text); }}
        40% {{ transform: translateY(-3px) scale(1.1); color: #FF3B30; text-shadow: 0 4px 15px rgba(255, 59, 48, 0.3); }}
        70% {{ transform: translateY(1px) scale(0.98); color: var(--dj-red); }}
        100% {{ opacity: 1; transform: translateY(0) scale(1); color: var(--dj-red); text-shadow: none; }}
    }}
    .dynamic-highlight {{
        display: inline-block; color: var(--dj-red); font-weight: 900;
        animation: textPop 0.85s cubic-bezier(0.34, 1.56, 0.64, 1) forwards;
    }}

    /* 🎯 1. 彻底居中的头部排版 */
    .elegant-header {{
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        gap: 18px; padding: 20px 0 40px 0; margin-bottom: 10px; text-align: center;
        animation: elegantEntrance 1s cubic-bezier(0.2, 0.8, 0.2, 1) forwards;
    }}
    .school-logo {{
        width: 130px; height: 130px; border-radius: 50%; object-fit: cover;
        box-shadow: 0 10px 25px rgba(139, 28, 49, 0.15); border: 4px solid var(--dj-bg); 
    }}
    .header-text-container {{ display: flex; flex-direction: column; align-items: center; }} 
    .sub-title-tag {{
        font-size: 13px; color: var(--dj-red); background-color: var(--dj-red-light);
        padding: 6px 18px; border-radius: 20px; font-weight: 700; letter-spacing: 1px; margin-bottom: 12px;
    }}
    .main-title {{ font-size: 40px; font-weight: 900; color: var(--dj-text); margin: 0; line-height: 1.2; letter-spacing: -1px; }}
    
    /* 🎯 2. 高级感 Segmented Control */
    div[role="radiogroup"] {{
        display: flex !important; width: 100% !important; flex-direction: row !important;
        background-color: var(--dj-red-light) !important; 
        padding: 5px !important; border-radius: 14px !important; gap: 4px !important;
    }}
    div[role="radiogroup"] label input[type="radio"] {{ display: none !important; }}
    div[role="radiogroup"] label input[type="radio"] + * {{ display: none !important; }}
    
    div[role="radiogroup"] label {{
        padding: 12px 4px !important; border-radius: 10px !important; margin: 0 !important;
        background-color: transparent !important; transition: all 0.35s cubic-bezier(0.25, 0.8, 0.25, 1) !important;
        cursor: pointer; display: flex !important; align-items: center !important; justify-content: center !important; flex: 1; 
    }}
    div[role="radiogroup"] label p {{ font-size: 16px !important; font-weight: 600 !important; color: var(--dj-text); margin: 0 !important; opacity: 0.6; transition: all 0.3s ease; }}
    
    div[role="radiogroup"] label:has(input:checked) {{ 
        background-color: #ffffff !important; 
        box-shadow: 0 3px 10px rgba(0,0,0,0.08), 0 1px 3px rgba(0,0,0,0.04) !important; 
        transform: scale(1); z-index: 1; 
    }}
    @media (prefers-color-scheme: dark) {{
        div[role="radiogroup"] label:has(input:checked) {{ background-color: #2c2c2e !important; }}
    }}
    
    div[role="radiogroup"] label:has(input:checked) p {{ color: var(--dj-red) !important; font-weight: 800 !important; opacity: 1; }}
    div[role="radiogroup"] label:active {{ transform: scale(0.94) !important; opacity: 0.8 !important; transition: all 0.1s ease !important; }}
    
    /* 隐藏无用的状态锚点，防止按钮上方出现多余空隙 */
    div[data-testid="column"] div[data-testid="element-container"]:has(.btn-status) {{ 
        display: none !important; 
        height: 0px !important;
        margin: 0px !important;
    }}

    /* 🎯 3. 班级按钮高级质感 (通用基础) */
    div.stButton > button {{
        border-radius: 14px !important; 
        font-weight: 700 !important; 
        transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important; 
        padding: 14px !important; 
        border: 1.5px solid transparent !important;
    }}
    div.stButton > button:active {{ transform: scale(0.92) translateY(2px) !important; filter: brightness(0.9) !important; transition: all 0.1s ease !important; }}

    /* ⚪ 未提交状态 (纯洁白/卡片底色) */
    div[data-testid="column"]:has(.btn-status-white) div.stButton > button {{ 
        background-color: var(--dj-card-bg) !important; 
        color: var(--dj-text) !important; 
        border-color: rgba(0,0,0,0.03) !important; 
    }}
    div[data-testid="column"]:has(.btn-status-white) div.stButton > button:hover {{ 
        transform: translateY(-3px) !important; 
        box-shadow: 0 8px 16px rgba(0,0,0,0.06) !important; 
        border-color: var(--dj-red-light) !important; 
    }}

    /* 🟢 提交成功 (纯正实心 iOS Green) */
    div[data-testid="column"]:has(.btn-status-green) div.stButton > button {{ 
        background-color: #34C759 !important; 
        color: #ffffff !important; 
        box-shadow: 0 4px 12px rgba(52, 199, 89, 0.3) !important; 
    }}

    /* 🔴 审查驳回 (纯正实心 iOS Red) */
    div[data-testid="column"]:has(.btn-status-red) div.stButton > button {{ 
        background-color: #FF3B30 !important; 
        color: #ffffff !important; 
        box-shadow: 0 4px 12px rgba(255, 59, 48, 0.3) !important; 
    }}

    /* 🟡 等待审核 (纯正实心 iOS Orange) */
    div[data-testid="column"]:has(.btn-status-pending) div.stButton > button {{ 
        background-color: #FF9500 !important; 
        color: #ffffff !important; 
        box-shadow: 0 4px 12px rgba(255, 149, 0, 0.3) !important; 
    }}
    
    @media (max-width: 768px) {{
        .school-logo {{ width: 110px; height: 110px; }} 
        .main-title {{ font-size: clamp(24px, 7vw, 30px); white-space: nowrap !important; letter-spacing: -0.5px; }}
        .block-container {{ padding-left: 0.8rem !important; padding-right: 0.8rem !important; }}
        div.stButton > button {{ padding: 14px 2px !important; font-size: 14px !important; border-radius: 12px !important; }}
        [data-testid="column"] {{ min-width: 30% !important; flex: 1 1 30% !important; padding: 0 6px !important; margin-bottom: 8px !important; }}
        div[role="radiogroup"] label {{ padding: 10px 4px !important; }}
        div[role="radiogroup"] label p {{ font-size: 15px !important; }}
    }}
    </style>
    
    <!-- 渲染动态头部容器 -->
    <div class="elegant-header">
        {logo_html}
        <div class="header-text-container">
            <div class="sub-title-tag" style="margin-bottom: 6px;">东江中学 · 青志联班级管理部</div>
            <div class="sub-title-tag">奉献 友爱 互助 进步</div>
            <h1 class="main-title">班级活动审查系统</h1>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    if not st.session_state.current_batch:
        st.info("📭 当前暂无正在收集的活动项目，请等待管理员发布。")
    else:
        activity_type = st.radio("📌 请选择活动类型", ["志愿服务活动", "团日活动"], horizontal=True)
        clean_act_type = "志愿服务活动" if "志愿" in activity_type else "团日活动"
        st.divider()
        
        grade_choice = st.radio("🎓 选择年级", ["高一", "高二", "高三"], horizontal=True)
        st.markdown(f"### 请选择 <span class='dynamic-highlight'>{grade_choice}</span> 的班级", unsafe_allow_html=True)
        
        current_db_status = get_class_status_from_db(st.session_state.current_batch, clean_act_type, grade_choice)
        total_classes = st.session_state.grade_config[grade_choice]
        
        # 🌟 核心革新：定义弹出窗口函数 (Dialog 弹窗)
        @st.dialog("📤 班级活动文件自动审查与提交通道")
        def show_submission_dialog(clean_act_type, grade_choice, current_class, class_info):
            st.markdown(f"#### 当前操作：<span class='dynamic-highlight'>{grade_choice} {current_class}</span> - {clean_act_type}", unsafe_allow_html=True)
            
            # 👇 全新升级：带呼吸浮动效果 + 显眼边框的加大警告框 👇
            st.markdown("""
            <style>
            /* 警告框的上下浮动与阴影呼吸动画 */
            @keyframes floatWarning {
                0% { transform: translateY(0px); box-shadow: 0 2px 8px rgba(255, 59, 48, 0.15); }
                50% { transform: translateY(-4px); box-shadow: 0 8px 16px rgba(255, 59, 48, 0.4); }
                100% { transform: translateY(0px); box-shadow: 0 2px 8px rgba(255, 59, 48, 0.15); }
            }
            .floating-warning-box {
                color: #FF3B30;
                font-weight: 900;
                font-size: 17px; /* 字号加大到 17px */
                text-align: center;
                padding: 12px;
                margin-top: -10px;
                margin-bottom: 12px;
                border: 2px solid #FF3B30; /* 添加纯正的红色外框 */
                border-radius: 12px;
                background-color: rgba(255, 59, 48, 0.08); /* 浅红色警示底色 */
                animation: floatWarning 2s ease-in-out infinite; /* 无限循环的悬浮动画 */
                letter-spacing: 1px;
            }
            </style>
            <div class="floating-warning-box">
                ！！！文件上传次数只有三次，请检查无误后上传！！！
            </div>
            """, unsafe_allow_html=True)
            
            st.divider()
            
            # 🌟 核心革新：定义弹出窗口函数 (Dialog 弹窗)
        @st.dialog("📤 班级活动文件自动审查与提交通道")
        def show_submission_dialog(clean_act_type, grade_choice, current_class, class_info):
            st.markdown(f"#### 当前操作：<span class='dynamic-highlight'>{grade_choice} {current_class}</span> - {clean_act_type}", unsafe_allow_html=True)
            
            # (这里是你之前添加的红色悬浮警告框代码)
            st.markdown("""
            <style>
            /* 警告框的上下浮动与阴影呼吸动画 */
            ... (略，保持原有代码) ...
            """, unsafe_allow_html=True)
            
            st.divider()

            # 👇 新增：管理员专属控制台 (重置次数 + 强制通过) 👇
            if st.session_state.get('admin_logged_in', False):
                st.info("🛠️ **管理员特权操作台**")
                
                col_admin1, col_admin2 = st.columns(2)
                
                # 功能 1：重置次数
                if col_admin1.button("🔄 次数归零", use_container_width=True):
                    sb.table("submissions").update({"attempts": 0}).match({
                        "batch_name": st.session_state.current_batch,
                        "activity_type": clean_act_type,
                        "grade": grade_choice,
                        "class_name": current_class
                    }).execute()
                    get_class_status_from_db.clear()
                    st.rerun()
                    
                # 功能 2：强制绿灯通过
                if col_admin2.button("✅ 强制通过", use_container_width=True):
                    # 先尝试更新现有记录，如果没有（未提交状态），则插入一条纯净的绿灯空记录
                    res = sb.table("submissions").select("id").match({
                        "batch_name": st.session_state.current_batch,
                        "activity_type": clean_act_type,
                        "grade": grade_choice,
                        "class_name": current_class
                    }).execute()
                    
                    if len(res.data) > 0:
                        # 记录已存在（可能被驳回或待审），直接改状态
                        sb.table("submissions").update({"status": "green"}).match({
                            "batch_name": st.session_state.current_batch,
                            "activity_type": clean_act_type,
                            "grade": grade_choice,
                            "class_name": current_class
                        }).execute()
                    else:
                        # 记录完全不存在（白板状态），生成一条“特批空文件”绿灯记录
                        sb.table("submissions").insert({
                            "batch_name": st.session_state.current_batch,
                            "activity_type": clean_act_type,
                            "grade": grade_choice,
                            "class_name": current_class,
                            "status": "green",
                            "attempts": 0,
                            "images_data": [],  # 空图片
                            "file_path": "",    # 故意留空路径，打包时触发防错机制跳过
                            "original_filename": f"{current_class}_特批免检.docx" 
                        }).execute()
                        
                    get_class_status_from_db.clear()
                    st.rerun()
                    
                st.divider()
            # 👆 新增结束 👆

                      
            if class_info["status"] == "green": 
                st.success("✅ 该班级已通过审查！无需再次提交。")
            elif class_info["status"] == "pending": 
                st.warning("⏳ 文件已成功提交，正在等待管理员后台人工审核图片...")
                st.info("🔒 审核期间暂时锁定上传通道。如被驳回，可再次提交。")
            elif class_info["attempts"] >= 3: 
                st.error("🚫 提交次数已耗尽，通道永久关闭。")
            else:
                st.info("📱 **手机端提交指引**：请先在 WPS 或微信中将填好的文档“另存为/保存到手机本地”，然后再点击下方按钮上传。")
                uploaded_file = st.file_uploader(f"上传《{clean_act_type}模板.docx》", type="docx")
                if uploaded_file is not None:
                    errors = []
                    try:
                        doc = Document(uploaded_file)
                        all_text = "".join([p.text for p in doc.paragraphs] + [cell.text for t in doc.tables for r in t.rows for cell in r.cells])
                        full_clean_text = re.sub(r'\s+', '', all_text).replace("：", ":").replace("、", "")
                        
                        if clean_act_type == "志愿服务活动":
                            participants_count = 0
                            time_duration_str = ""
                            nums = re.findall(r'参与人数.*?(\d+)', full_clean_text)
                            if nums: participants_count = int(nums[0])
                            
                            if "活动时长" in full_clean_text:
                                time_match = re.search(r'活动时长.{0,15}', full_clean_text)
                                if time_match: time_duration_str = time_match.group()
                                    
                            if participants_count == 0: errors.append("第2项错误：未读取到参与人数。")
                            elif grade_choice == "高一" and not (10 <= participants_count <= 12): errors.append(f"高一人数应为10-12人，实际 {participants_count} 人。")
                            elif grade_choice in ["高二", "高三"] and participants_count < 10: errors.append(f"{grade_choice}人数应≥10人，实际 {participants_count} 人。")
                            
                            phone_numbers = re.findall(r'1[3-9]\d{9}', full_clean_text)
                            if len(phone_numbers) != participants_count: errors.append(f"第4项错误：检测到的手机号数量({len(phone_numbers)}个)与填写的参与人数({participants_count}人)不一致。")
                                    
                            if not time_duration_str: errors.append("第3项错误：未读取到活动时长。")
                            else:
                                time_passed, time_msg = check_time_duration(time_duration_str)
                                if not time_passed: errors.append(f"第3项错误：{time_msg}")
                                
                            uploaded_file.seek(0)
                            images = extract_images_from_docx(uploaded_file)
                            if len(images) != 4: errors.append(f"第9项错误：图片必须为4张，实际提取到 {len(images)} 张。")
                            else:
                                with st.spinner("视觉引擎高速运转中... 正在进行多模态 AI 图像审查 (约5~10秒)"):
                                    ai_passed, ai_reason = review_images_with_ai(images)
                                if not ai_passed: errors.append(f"AI 图片审查未通过：{ai_reason}")
                                
                        elif clean_act_type == "团日活动":
                            uploaded_file.seek(0)
                            images = extract_images_from_docx(uploaded_file)
                            if len(images) != 3: errors.append(f"第九项错误：必须提交 3 张图片，实际 {len(images)} 张。")
                                
                            required_keywords = ["活动背景", "活动主题", "活动人数", "活动目的", "活动时间", "活动形式", "活动流程", "学生感想"]
                            for i, keyword in enumerate(required_keywords):
                                if keyword not in full_clean_text:
                                    errors.append(f"未找到【{keyword}】这一项标题。")
                                else:
                                    content_after = full_clean_text.split(keyword)[-1]
                                    content_between = content_after.split(required_keywords[i+1])[0] if i < len(required_keywords) - 1 else content_after
                                    pure_content = re.sub(r'[^\w\u4e00-\u9fa5]', '', content_between)
                                    if len(pure_content) < 2: errors.append(f"必填项【{keyword}】未填写有效内容。")

                        new_attempts = class_info["attempts"] + 1
                        uploaded_file.seek(0)
                        file_bytes = uploaded_file.read()
                        
                        sb.table("submissions").delete().match({
                            "batch_name": st.session_state.current_batch,
                            "activity_type": clean_act_type,
                            "grade": grade_choice,
                            "class_name": current_class
                        }).execute()

                        if len(errors) == 0:
                            final_status = "green" if clean_act_type == "志愿服务活动" else "pending"
                            images_b64 = [base64.b64encode(img).decode('utf-8') for img in images] if clean_act_type == "团日活动" else []
                            
                            import time
                            safe_storage_path = f"doc_{int(time.time() * 1000)}.docx"
                            
                            sb.storage.from_("school-docs").upload(
                                path=safe_storage_path, 
                                file=file_bytes, 
                                file_options={"upsert": "true"}
                            )
                            
                            sb.table("submissions").insert({
                                "batch_name": st.session_state.current_batch,
                                "activity_type": clean_act_type,
                                "grade": grade_choice,
                                "class_name": current_class,
                                "status": final_status,
                                "attempts": new_attempts,
                                "images_data": images_b64,
                                "file_path": safe_storage_path,
                                "original_filename": uploaded_file.name
                            }).execute()
                            
                            st.success("🎉 审查/提交成功！即将自动返回...")
                            time.sleep(2) # 留出2秒时间给用户看成功提示，随后自动关闭弹窗
                            st.rerun() 
                        else:
                            sb.table("submissions").insert({
                                "batch_name": st.session_state.current_batch,
                                "activity_type": clean_act_type,
                                "grade": grade_choice,
                                "class_name": current_class,
                                "status": "red",
                                "attempts": new_attempts
                            }).execute()
                            st.error("❌ 审查未通过：\n" + "\n".join([f"{i+1}. {err}" for i, err in enumerate(errors)]))
                            if st.button("确认并重试"):
                                st.rerun()
                                
                    except Exception as e:
                        st.error(f"解析出错，文档可能损坏: {e}")
                    finally:
                        get_class_status_from_db.clear()
                        if 'file_bytes' in locals(): del file_bytes
                        if 'images' in locals(): del images
                        if 'images_b64' in locals(): del images_b64
                        if 'all_text' in locals(): del all_text
                        import gc
                        gc.collect() 
        
        # 🌟 核心革新：全版本通杀的 CSS 变色法 (无视 Streamlit 底层更新)
        universal_css = """
        <style>
        /* 1. 彻底隐藏状态锚点，兼容新老 Streamlit DOM */
        div[data-testid="element-container"]:has(.btn-status),
        div[data-testid="stElementContainer"]:has(.btn-status),
        .element-container:has(.btn-status) {
            display: none !important; height: 0px !important; margin: 0px !important; padding: 0px !important;
        }

        /* 2. 班级按钮高级质感 */
        div.stButton > button {
            border-radius: 14px !important; font-weight: 700 !important; 
            transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important; padding: 14px !important; 
        }

        /* ⚪ 未提交 (卡片底色实心) */
        div[data-testid="element-container"]:has(.btn-status-white) + div[data-testid="element-container"] button,
        div[data-testid="stElementContainer"]:has(.btn-status-white) + div[data-testid="stElementContainer"] button {
            background-color: var(--dj-card-bg) !important; color: var(--dj-text) !important; border: 1.5px solid rgba(0,0,0,0.03) !important;
        }
        div[data-testid="element-container"]:has(.btn-status-white) + div[data-testid="element-container"] button:hover,
        div[data-testid="stElementContainer"]:has(.btn-status-white) + div[data-testid="stElementContainer"] button:hover {
            transform: translateY(-3px) !important; box-shadow: 0 8px 16px rgba(0,0,0,0.06) !important; border-color: var(--dj-red-light) !important;
        }

        /* 🟢 提交成功 (纯正实心 iOS Green) */
        div[data-testid="element-container"]:has(.btn-status-green) + div[data-testid="element-container"] button,
        div[data-testid="stElementContainer"]:has(.btn-status-green) + div[data-testid="stElementContainer"] button {
            background-color: #34C759 !important; color: #ffffff !important; border: none !important; box-shadow: 0 4px 12px rgba(52, 199, 89, 0.3) !important;
        }

        /* 🔴 审查驳回 (纯正实心 iOS Red) */
        div[data-testid="element-container"]:has(.btn-status-red) + div[data-testid="element-container"] button,
        div[data-testid="stElementContainer"]:has(.btn-status-red) + div[data-testid="stElementContainer"] button {
            background-color: #FF3B30 !important; color: #ffffff !important; border: none !important; box-shadow: 0 4px 12px rgba(255, 59, 48, 0.3) !important;
        }

        /* 🟡 等待审核 (纯正实心 iOS Orange) */
        div[data-testid="element-container"]:has(.btn-status-pending) + div[data-testid="element-container"] button,
        div[data-testid="stElementContainer"]:has(.btn-status-pending) + div[data-testid="stElementContainer"] button {
            background-color: #FF9500 !important; color: #ffffff !important; border: none !important; box-shadow: 0 4px 12px rgba(255, 149, 0, 0.3) !important;
        }
        </style>
        """
        st.markdown(universal_css, unsafe_allow_html=True)
        
        # 渲染班级方块（点击直接召唤中央弹窗）
        for i in range(1, total_classes + 1, 6):
            cols = st.columns(6)
            for j in range(6):
                class_idx = i + j
                if class_idx <= total_classes:
                    cls_name = f"{class_idx}班"
                    data = current_db_status.get(cls_name, {"status": "white", "attempts": 0})
                    
                    col = cols[j]
                    
                    # 在按钮正上方埋入隐形锚点，CSS 将利用 "+" 兄弟选择器精准命中它下方的按钮
                    col.markdown(f"<div class='btn-status btn-status-{data['status']}'></div>", unsafe_allow_html=True)
                    
                    btn_label = f"{cls_name}\n({data['attempts']}/3)"
                    # 点击操作直接触发弹窗函数！
                    if col.button(btn_label, key=f"btn_{clean_act_type}_{grade_choice}_{cls_name}", use_container_width=True):
                        show_submission_dialog(clean_act_type, grade_choice, cls_name, data)
        # ================= 全局网页最底部声明 =================
st.divider()
st.markdown("""
<div style="text-align: center; color: var(--dj-text); opacity: 0.45; font-size: 12px; padding: 20px 0 40px 0;">
    本网站由 邱世豪 与 Google Gemini 协作开发 | 出现问题欢迎积极反馈
</div>
""", unsafe_allow_html=True)