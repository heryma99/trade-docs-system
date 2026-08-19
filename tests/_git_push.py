import subprocess, os

ROOT = r'D:\WB文件\2026-07-30-09-36-10\贸易单证系统'
def run(cmd, cwd=ROOT, timeout=600):
    r = subprocess.run(cmd, shell=True, cwd=cwd, capture_output=True, text=True, encoding='utf-8', errors='ignore', timeout=timeout)
    return r.returncode, (r.stdout + r.stderr).strip()

rc, out = run('git commit -q -m "v1.5.44 清空旧图库(全部占位错图444张) 从JW PEI G Unit抽153张SPU真图+SPU前缀降级"')
print('commit:', out[:200] if rc else 'OK')

# 验证 commit 里有 153 张图
rc, out = run('git show --stat HEAD | grep -c "images/sku_thumb/"')
print('commit 图片数:', out.strip())

# push (force 到 main，全新仓库覆盖历史)
rc, out = run('git push --force origin main', timeout=600)
print('push:', out[-400:] if out else 'OK(无输出)')
print('EXIT:', rc)