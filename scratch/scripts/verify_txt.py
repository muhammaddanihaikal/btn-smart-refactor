import os, re

def verify_all_txt(base_dir):
    long_prefix = '\\\\?\\'
    abs_base = long_prefix + os.path.abspath(base_dir)
    
    total = 0
    issues = []
    
    for root, dirs, files in os.walk(abs_base):
        for f in files:
            if f.endswith('.txt') and not f.startswith('desktop'):
                total += 1
                txt_path = os.path.join(root, f)
                with open(txt_path, 'r', encoding='utf-8', errors='ignore') as fp:
                    content = fp.read()
                
                # Check STEPS
                steps_match = re.search(r'STEPS[^\n]*\n=+\n(.*?)\n=+', content, re.DOTALL)
                exp_match = re.search(r'EXPECTED RESULTS[^\n]*\n=+\n(.*)', content, re.DOTALL)
                
                steps_text = steps_match.group(1).strip() if steps_match else ''
                exp_text = exp_match.group(1).strip() if exp_match else ''
                
                has_step_num = bool(re.search(r'^\s*1\.', steps_text, re.MULTILINE))
                has_exp_num = bool(re.search(r'^\s*1\.', exp_text, re.MULTILINE))
                has_dash_bullet = bool(re.search(r'^\s*-\s+', steps_text, re.MULTILINE) or re.search(r'^\s*-\s+', exp_text, re.MULTILINE))
                has_inline_dashes = (' - ' in steps_text) or (' - ' in exp_text)
                
                if not has_step_num or not has_exp_num or has_dash_bullet or has_inline_dashes:
                    issues.append({
                        'file': f,
                        'has_step_num': has_step_num,
                        'has_exp_num': has_exp_num,
                        'has_dash_bullet': has_dash_bullet,
                        'has_inline_dashes': has_inline_dashes,
                        'steps': steps_text[:120],
                        'exp': exp_text[:120]
                    })

    print(f"Directory: {base_dir}")
    print(f"Total .txt files checked: {total}")
    print(f"Files with issues: {len(issues)}")
    if issues:
        print("\n=== SAMPLE ISSUES FOUND ===")
        for iss in issues[:15]:
            fname = iss['file']
            s_num = iss['has_step_num']
            e_num = iss['has_exp_num']
            d_bul = iss['has_dash_bullet']
            d_inl = iss['has_inline_dashes']
            st = iss['steps'].replace('\n', ' ')
            ex = iss['exp'].replace('\n', ' ')
            print(f"- File: {fname}")
            print(f"  StepNum={s_num} | ExpNum={e_num} | DashBullet={d_bul} | InlineDash={d_inl}")
            print(f"  Steps: {st}")
            print(f"  Exp:   {ex}\n")

verify_all_txt(r'D:\Project\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile')
verify_all_txt(r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Screenshot\Mobile')
