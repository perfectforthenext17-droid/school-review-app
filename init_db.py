import sqlite3

print("🔌 正在创建并连接本地 SQLite 数据库...")

try:
    # 1. 连接到本地数据库文件（如果文件不存在，它会自动帮你建一个）
    conn = sqlite3.connect('school_review.db')
    cursor = conn.cursor()

    # 2. 创建核心数据表 (对应你刚才在网页上建的那些列)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS submissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        batch_name TEXT,
        activity_type TEXT,
        grade TEXT,
        class_name TEXT,
        status TEXT,
        attempts INTEGER,
        images_data TEXT
    )
    ''')

    # 3. 尝试插入一条测试数据
    cursor.execute('''
    INSERT INTO submissions (batch_name, activity_type, grade, class_name, status, attempts)
    VALUES (?, ?, ?, ?, ?, ?)
    ''', ('系统测试批次', '系统启动测试', '高一', '测试班', 'white', 0))

    # 4. 保存修改并关闭
    conn.commit()
    conn.close()
    
    print("✅ 完美搞定！本地数据库已成功建立并写入测试数据！")
    print("📁 请查看你的左侧文件夹，是不是多了一个名叫 'school_review.db' 的文件？")

except Exception as e:
    print(f"❌ 出现意外错误: {e}")