from zhipuai import ZhipuAI

# ⚠️ 请务必把这里换成你真实的智谱 API Key
API_KEY = "ac67cb8c61804fe6bdcecdc13e8aee0a.v7IK0Kxs4EMLsMbr"

print("正在尝试连接智谱服务器，请稍候...")

try:
    client = ZhipuAI(api_key=API_KEY)
    # 先用纯文本模型 (glm-4) 测试最基本的网络连通性
    response = client.chat.completions.create(
        model="glm-4",
        messages=[{"role": "user", "content": "你好，请回复我收到。"}],
        timeout=15  # 强制设置 15 秒超时，防止无限卡死
    )
    print("✅ 连接完全成功！AI 回复：", response.choices[0].message.content)

except Exception as e:
    print("❌ 连接失败！系统报出的真实错误是：", e)
