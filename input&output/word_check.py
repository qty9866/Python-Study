import argparse
parser = argparse.ArgumentParser(description="用于演示参数的处理")
# -number 代表参数为可选参数
parser.add_argument("-number",help="输入的数字")
parser.add_argument("-string",help="输入的字符")
args = parser.parse_args()
print(f"你输入的数字是{args.number}")
print(f"你输入的字符是{args.string}")