import hashlib
import uuid
import json

# 存储用户信息和充值密文
users = {}
recharge_codes = {}


def generate_user(username):
    # 生成随机密码
    password = uuid.uuid4().hex[:8]
    # 生成用户ID
    user_id = uuid.uuid4().hex
    # 存储用户信息
    users[user_id] = {
        'username': username,
        'password': password,
        'points': 0
    }
    return user_id, password


def generate_recharge_code(user_id, points):
    # 生成唯一充值码
    recharge_code = hashlib.sha256(f"{user_id}{uuid.uuid4()}".encode()).hexdigest()
    recharge_codes[recharge_code] = {
        'user_id': user_id,
        'points': points,
        'used': False
    }
    return recharge_code


def save_data():
    # 保存用户信息和充值码到文件
    with open('users.json', 'w') as f:
        json.dump(users, f)
    with open('recharge_codes.json', 'w') as f:
        json.dump(recharge_codes, f)


def load_data():
    global users, recharge_codes
    try:
        with open('users.json', 'r') as f:
            users = json.load(f)
    except FileNotFoundError:
        users = {}
    try:
        with open('recharge_codes.json', 'r') as f:
            recharge_codes = json.load(f)
    except FileNotFoundError:
        recharge_codes = {}


def view_all_accounts():
    if users:
        print("\n当前所有用户账户信息:")
        for user_id, user_info in users.items():
            print(
                f"用户ID: {user_id}, 用户名: {user_info['username']}, 密码: {user_info['password']}, 积分: {user_info['points']}")
    else:
        print("没有账户信息。")


def main():
    load_data()

    while True:
        print("\n1. 生成用户账号")
        print("2. 生成充值密文")
        print("3. 查看所有账户")
        print("4. 退出")
        choice = input("请输入你的选择: ")

        if choice == "1":
            username = input("请输入用户名: ")
            user_id, password = generate_user(username)
            print(f"生成的用户信息 - 用户名: {username}, 密码: {password}, 用户ID: {user_id}")
            save_data()
        elif choice == "2":
            user_id = input("请输入用户ID以生成充值密文: ")
            if user_id in users:
                points = int(input("请输入充值积分: "))
                recharge_code = generate_recharge_code(user_id, points)
                print(f"生成的充值密文: {recharge_code}")
                save_data()
            else:
                print("无效的用户ID。")
        elif choice == "3":
            view_all_accounts()
        elif choice == "4":
            break
        else:
            print("无效的选择。")


if __name__ == "__main__":
    main()
