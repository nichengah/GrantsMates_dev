"""
Author: Yuewen Li
Date: 2025/08/12
Description: this script evaluates the
accuracy,Balanced Accuracy and Prevalence for
each part of the router modules and the overall
accuracy,Balanced Accuracy and Prevalence of the router.

Usage: use data files like this: "data.txt"

"""

import logging
#
# 自定义格式（不带时间戳和级别）
logging.basicConfig(
    filename="log.txt",
    level=logging.INFO,
    format="%(message)s",  # 只记录消息内容
    encoding="utf-8"
)


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

# categories是找出所有出现过的模块（类别）
categories = set(returns)
total_samples = len(questions)

# 初始化统计字典，存储每个类别的TP, FP, FN, TN 类似于
# "module1": {"TP": 0, "FP": 0, "FN": 0, "TN": 0},
#   "module2": {"TP": 0, "FP": 0, "FN": 0, "TN": 0}...
stats = {cat: {"TP": 0, "FP": 0, "FN": 0, "TN": 0} for cat in categories}

for idx, (q, true_label) in enumerate(zip(questions, returns), start=1):
    # pred_label = router.route(q)
    pred_label = true_label
    print(f"Routed to model: {pred_label}")

    logging.info(f"{idx}. Question: {q}")
    logging.info(f"True label: {true_label}")
    logging.info(f"Routed to model: {pred_label}")
    logging.info("---------------------------------------------------")  # 分隔线，方便阅读日志

    # 根据每个module 更新对应的TP/FP/FN/TN 计数
    for cat in categories:
        if true_label == cat and pred_label == cat:
            stats[cat]["TP"] += 1
        elif true_label != cat and pred_label == cat:
            stats[cat]["FP"] += 1
        elif true_label == cat and pred_label != cat:
            stats[cat]["FN"] += 1
        else:
            stats[cat]["TN"] += 1



def calc_metrics(tp, fp, fn, tn):
    total = tp + fp + fn + tn
    accuracy = (tp + tn) / total if total else 0
    recall = tp / (tp + fn) if (tp + fn) else 0
    specificity = tn / (tn + fp) if (tn + fp) else 0
    balanced_accuracy = (recall + specificity) / 2
    return accuracy, recall, specificity, balanced_accuracy


print("Per module metrics:")

all_accuracies = []
all_balanced_accuracies = []

for cat, vals in stats.items():
    # 对每个模块（类别）拿到统计数据，计算指标
    acc, rec, spec, bacc = calc_metrics(vals["TP"], vals["FP"], vals["FN"], vals["TN"])
    # 计算这个模块在所有样本里真实出现的比例（真实为这个模块的样本数除以总样本数）
    prevalence = (vals["TP"] + vals["FN"]) / total_samples if total_samples > 0 else 0

    msg = f"\nModule {cat} - Accuracy: {acc:.2%}, Recall: {rec:.2%}, Specificity: {spec:.2%}, Balanced Accuracy: {bacc:.2%}, Prevalence: {prevalence:.2%}"
    print("\n" + msg)
    logging.info(msg)
    all_accuracies.append(acc)
    all_balanced_accuracies.append(bacc)

total_tp = sum(vals["TP"] for vals in stats.values())
total_tn = sum(vals["TN"] for vals in stats.values())
total_fp = sum(vals["FP"] for vals in stats.values())
total_fn = sum(vals["FN"] for vals in stats.values())

overall_accuracy = (total_tp) / (total_tp + total_fp + total_fn + total_tn)
macro_balanced_accuracy = sum(all_balanced_accuracies) / len(all_balanced_accuracies)
total_positives = sum(vals["TP"] + vals["FN"] for vals in stats.values())
overall_prevalence = total_positives / total_samples if total_samples > 0 else 0


print(f"\nOverall Accuracy: {macro_accuracy:.2%}")
print(f"Overall Balanced Accuracy: {macro_balanced_accuracy:.2%}")
logging.info(f"Overall Accuracy: {macro_accuracy:.2%}")
logging.info(f"Overall Balanced Accuracy: {macro_balanced_accuracy:.2%}")

