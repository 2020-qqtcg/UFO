import json
from collections import defaultdict

input_file="./outputs/eval_custom_custom.json"
output_file="./outputs/eval_result_analyse.json"

# 读取数据
with open(input_file, 'r', encoding='utf-8') as f:
    data = json.load(f)

# 创建统计字典
soft_stats = defaultdict(list)
hard_stats = defaultdict(list)

# 统计每个 soft/hard_restriction 值对应的 id 列表
for item in data:
    soft = item.get("soft_restriction")
    hard = item.get("hard_restriction")
    id_ = item.get("id")

    soft_stats[soft].append(id_)
    hard_stats[hard].append(id_)

# 构造输出结构
output = {
    "soft_restriction": {
        str(k): {
            "count": len(v),
            "ids": v
        } for k, v in soft_stats.items()
    },
    "hard_restriction": {
        str(k): {
            "count": len(v),
            "ids": v
        } for k, v in hard_stats.items()
    }
}

# 输出统计结果
print("Soft Restriction 分布:")
for k, v in soft_stats.items():
    print(f"  soft_restriction = {k}, {len(v)} 个")

print("\nHard Restriction 分布:")
for k, v in hard_stats.items():
    print(f"  hard_restriction = {k}, {len(v)} 个")


# 写入到新文件
with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(output, f, indent=4, ensure_ascii=False)

print(f"统计结果已写入{output_file}")
