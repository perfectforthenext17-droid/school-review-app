# [📍 修改标记：配置数据库]
from supabase import create_client, Client

# 把下面两个字符串替换成你刚刚在网页上复制的真实 URL 和 Key
SUPABASE_URL = "https://你的项目前缀.supabase.co"
SUPABASE_KEY = "你的_ANON_KEY"

print("🔌 正在尝试连接底层数据库...")

try:
    # 建立连接
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    
    # 尝试向刚建好的 submissions 表里插入一条测试数据
    test_data = {
        "batch_name": "系统测试批次",
        "activity_type": "系统启动测试",
        "grade": "高一",
        "class_name": "测试班",
        "status": "white",
        "attempts": 0
    }
    
    data, count = supabase.table("submissions").insert(test_data).execute()
    
    print("✅ 数据库打通成功！数据已成功写入云端！")
    print("返回的数据记录是:", data)
    
except Exception as e:
    print(f"❌ 数据库连接失败，错误详情: {e}")