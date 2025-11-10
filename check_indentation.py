# 检查Python文件的缩进是否一致（4个空格）

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

print("检查缩进不一致的行:")
print("="*50)

indentation_issues = []

for i, line in enumerate(lines, 1):
    stripped = line.lstrip()
    if stripped and not stripped.startswith('#'):
        leading_spaces = len(line) - len(stripped)
        # 检查缩进是否是4的倍数
        if leading_spaces > 0 and leading_spaces % 4 != 0:
            indentation_issues.append((i, leading_spaces, line.rstrip()))

if indentation_issues:
    for line_num, spaces, content in indentation_issues:
        print(f"第{line_num}行: {spaces}个空格缩进 - {content}")
else:
    print("没有发现缩进不一致的问题！")

print("\n检查空行后的缩进:")
print("="*50)

# 检查空行后的缩进是否正确
prev_empty = False
for i, line in enumerate(lines, 1):
    if not line.strip():
        prev_empty = True
    elif prev_empty and line.strip() and not line.strip().startswith('#'):
        leading_spaces = len(line) - len(line.lstrip())
        if leading_spaces % 4 != 0:
            print(f"空行后的第{line_num}行: {spaces}个空格缩进 - {content}")
        prev_empty = False