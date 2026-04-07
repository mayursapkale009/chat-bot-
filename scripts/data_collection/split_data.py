import os
import json
import random

def split_jsonl(input_file, train_file, val_file, test_file, split_ratios=(0.7, 0.1, 0.2)):
    if not os.path.exists(input_file):
        print(f"Skipping split: {input_file} not found.")
        return

    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    random.shuffle(lines)
    
    total = len(lines)
    train_end = int(total * split_ratios[0])
    val_end = train_end + int(total * split_ratios[1])
    
    train_lines = lines[:train_end]
    val_lines = lines[train_end:val_end]
    test_lines = lines[val_end:]
    
    with open(train_file, 'w', encoding='utf-8') as f:
        f.writelines(train_lines)
    with open(val_file, 'w', encoding='utf-8') as f:
        f.writelines(val_lines)
    with open(test_file, 'w', encoding='utf-8') as f:
        f.writelines(test_lines)
        
    print(f"Successfully split {total} lines into Train({len(train_lines)}), Val({len(val_lines)}), Test({len(test_lines)}).")

if __name__ == "__main__":
    base_dir = os.path.join(os.path.dirname(__file__), '../../data/processed/intent/hindi')
    split_jsonl(
        os.path.join(base_dir, 'all.jsonl'),
        os.path.join(base_dir, 'train.jsonl'),
        os.path.join(base_dir, 'val.jsonl'),
        os.path.join(base_dir, 'test.jsonl')
    )
