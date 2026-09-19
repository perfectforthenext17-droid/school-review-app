import time
from supabase import create_client

# 替换为你刚才复制的实际值
SUPABASE_URL = "https://nskalqccelbtrwjlsmyh.supabase.co"
SUPABASE_KEY = "sb_publishable_b6htTHXki7wLrIEr2vwjaQ_OlXATkLJ"

print("🔍 正在测试国内网络直连 Supabase（请确保已关闭 VPN）...")
start_time = time.time()

try:
    client = create_client(SUPABASE_URL, SUPABASE_KEY)
    # 测试基础元数据读取
    res = client.table("batches").select("*").limit(1).execute()
    elapsed = round(time.time() - start_time, 2)
    print(f"✅ 连接完全成功！耗时: {elapsed} 秒，网络顺畅。")
except Exception as e:
    elapsed = round(time.time() - start_time, 2)
    print(f"❌ 连接测试失败（耗时 {elapsed} 秒），错误信息: {e}")