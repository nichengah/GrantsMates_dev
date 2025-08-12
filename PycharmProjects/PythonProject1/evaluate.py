questions = []
returns = []

with open("data.txt", "r", encoding="utf-8") as f:
    lines = [line.strip() for line in f if line.strip()]  # 去空行，去换行符

for i in range(0, len(lines), 2):  # 每两行处理一组
    q_line = lines[i]
    r_line = lines[i+1] if i+1 < len(lines) else ""

    # 去除前缀，得到纯文本(two parts questions and returns)
    q_text = q_line[len("Q: "):] if q_line.startswith("Q: ") else q_line
    r_text = r_line[len("return: "):] if r_line.startswith("return: ") else r_line

    questions.append(q_text)
    returns.append(r_text)

    # def router

    # 初始化统计字典
stats = {
        "module1": {"correct": 0, "total": 0},
        "module2": {"correct": 0, "total": 0},
        "module3": {"correct": 0, "total": 0},
        "module4": {"correct": 0, "total": 0},
}

    # total_count = 0
    # correct_count = 0
for q, expected in zip(questions, returns):
    # 得到router的答案
    answer,module = router(q)
    is_correct = (answer == expected)
    print(f"Q: {q}")
    print(f"Expected: {expected}")
    print(f"Router answer: {answer}")
    print(f"Module: {module}")
    print("Match:", is_correct)
    print("---")

    stats[module]["total"] += 1
    if is_correct:
        stats[module]["correct"] += 1


# accuracy = correct_count / total_count
# 统计每个模块准确率的代码
for module, data in stats.items():
    total = data["total"]
    correct = data["correct"]
    accuracy = correct / total if total > 0 else 0
    # accuracy of each part
    print(f"Module {module} - Total: {total}, Correct: {correct}, Accuracy: {accuracy:.2%}")


# 计算整体准确率
total_correct = sum(data["correct"] for data in stats.values())
total_questions = sum(data["total"] for data in stats.values())
overall_accuracy = total_correct / total_questions if total_questions > 0 else 0

print(f"Overall accuracy: {overall_accuracy:.2%}")


# # 测试打印
# for q, r in zip(questions, returns):
#     print("Q:", q)
#     print("Return:", r)
#     print("---")
