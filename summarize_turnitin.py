import json

with open('turnitin_extracted_matches.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print(f"{'Page':<6} | {'Words':<6} | {'Sample Text'}")
print("-" * 80)
total_words = 0
for p, item in sorted(data.items(), key=lambda x: int(x[0])):
    words = item['count']
    total_words += words
    sample = item['text'][:75].replace('\n', ' ')
    print(f"{p:<6} | {words:<6} | {sample}")

print("-" * 80)
print(f"Total highlighted words detected: {total_words}")
