import subprocess, os, sys

ROOT = r'D:\WB文件\2026-07-30-09-36-10\贸易单证系统'
GIT = r'C:\Users\cn\.workbuddy\binaries\node\versions\22.22.2\node.exe'  # 不用

# 用系统 git
def run(cmd, cwd=ROOT):
    r = subprocess.run(cmd, shell=True, cwd=cwd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
    return r.returncode, (r.stdout + r.stderr).strip()

# 1) 备份并删除损坏 .git
gitdir = os.path.join(ROOT, '.git')
if os.path.exists(gitdir):
    os.rename(gitdir, gitdir + '_broken_bak')
    print('已备份损坏 .git -> .git_broken_bak')

# 2) 确保 .gitignore 存在
gi = os.path.join(ROOT, '.gitignore')
if not os.path.exists(gi):
    with open(gi, 'w', encoding='utf-8') as f:
        f.write('tests/_out/\nimages/products_sku/\nimages/products/\nimages/_jwpei_ok/\nimages/sku_all/\ntds_dist/\nnode_modules/\n*.log\n.git_broken_bak/\n')
    print('已写 .gitignore')

# 3) 全新 init + 提交图片和关键文件
rc, out = run('git init -b main')
print('init:', out[:80])
rc, out = run('git config user.email "deploy@local" && git config user.name "deploy"')
print('config:', out[:50])
# token 写入 remote
TOKEN = open(r'C:\Users\cn\.workbuddy\connectors\3fe83c35-d7d3-4e71-869e-097580283ed4\tokens\github.txt', 'r', encoding='utf-8').read().strip()
rc, out = run(f'git remote add origin https://x-access-token:{TOKEN}@github.com/heryma99/trade-docs-system.git')
print('remote:', out[:60])

# 4) add 图片目录 + 关键 js
rc, out = run('git add images/sku_thumb js/sku_image_index.js js/engine.js js/ui.js index.html js/product_image_map.js userdata.json')
print('add:', out[:100])
rc, out = run('git status --short | head -20')
print('status 前20行:\n', out[:800])