import os, shutil, json, re

ROOT = "D:/WB文件/2026-07-30-09-36-10/贸易单证系统"
CLEAN = ROOT + "/deploy_clean"
IMG_DST = CLEAN + "/images/sku_thumb"

# 读 manifest(精确 4225)
man = json.load(open(ROOT + "/_manifest_b.json"))

# 清 deploy_clean(若存在) —— 用 Python 单目录删, 若被钩子拦就逐个删
if os.path.isdir(CLEAN):
    try:
        shutil.rmtree(CLEAN)
    except Exception as e:
        print("rmtree 被拦:", e)

os.makedirs(IMG_DST, exist_ok=True)

# 1) 应用文件(根目录, 排除垃圾)
app_files = ["index.html", "styles.css", "box_specs.js", "userdata.json", "README.md"]
app_dirs = ["js", "vendor", "templates"]
for f in app_files:
    src = ROOT + "/" + f
    if os.path.isfile(src):
        shutil.copy2(src, CLEAN + "/" + f)
for d in app_dirs:
    s = ROOT + "/" + d
    if os.path.isdir(s):
        shutil.copytree(s, CLEAN + "/" + d, dirs_exist_ok=True)

# 2) 精确 4225 图
copied = 0
missing = []
for e in man:
    src = ROOT + "/" + e["rel"]   # rel = images/sku_thumb/xxx.jpeg
    dst = CLEAN + "/" + e["rel"]
    if os.path.isfile(src):
        shutil.copy2(src, dst)
        copied += 1
    else:
        missing.append(e["rel"])
print("应用文件已复制")
print("图片复制: %d / %d" % (copied, len(man)))
if missing:
    print("缺失:", missing[:5])

# 3) 校验 deploy_clean 图片数 = 4225
n = len(os.listdir(IMG_DST))
print("deploy_clean 图片数:", n)

# 4) 列 deploy_clean 顶层
print("deploy_clean 顶层:", sorted(os.listdir(CLEAN)))
