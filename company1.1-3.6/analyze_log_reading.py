"""分析 log 讀取流程，檢查產業別是否成功抓取"""
import json
from pathlib import Path
from env_log_reader import load_latest_environment_log, get_prompt_context

print("="*80)
print("=== 分析 log 讀取流程 ===")
print("="*80)

# Step 1: 檢查原始 log 文件
print("\n【Step 1】檢查原始 log 文件中的產業別")
print("-"*80)
log_dir = Path(r"C:\Users\User\Desktop\ESG_Output\_Backend\user_logs")
files = sorted(log_dir.glob("*.json"), key=lambda f: f.stat().st_mtime, reverse=True)

print(f"找到 {len(files)} 個 log 文件\n")

# 檢查最新的文件
latest_file = files[0] if files else None
if latest_file:
    print(f"最新文件: {latest_file.name}")
    try:
        with open(latest_file, 'r', encoding='utf-8') as f:
            latest_data = json.load(f)
        print(f"  包含 industry key: {'industry' in latest_data}")
        print(f"  industry 值: {repr(latest_data.get('industry', 'NOT FOUND'))}")
        print(f"  step: {latest_data.get('step', 'N/A')}")
    except Exception as e:
        print(f"  讀取錯誤: {e}")

# 檢查包含 industry 的文件
print(f"\n包含 industry 的文件:")
industry_files = []
for f in files[:10]:
    try:
        data = json.load(open(f, encoding='utf-8'))
        if 'industry' in data and data.get('industry'):
            industry_files.append((f, data.get('industry')))
            print(f"  - {f.name}: {repr(data.get('industry'))}")
    except:
        pass

# Step 2: 測試 load_latest_environment_log
print("\n" + "="*80)
print("【Step 2】測試 load_latest_environment_log()")
print("-"*80)
log_data = load_latest_environment_log()

if log_data:
    industry = log_data.get("industry", "")
    print(f"✅ log_data 成功載入")
    print(f"   industry: {repr(industry)}")
    print(f"   industry 是否為空: {industry == ''}")
    print(f"   industry 是否為 None: {industry is None}")
    print(f"   industry 長度: {len(industry) if industry else 0}")
    print(f"   industry 類型: {type(industry)}")
else:
    print("❌ log_data 為 None")

# Step 3: 測試 get_prompt_context
print("\n" + "="*80)
print("【Step 3】測試 get_prompt_context()")
print("-"*80)
context = get_prompt_context(log_data)

if context:
    industry = context.get("industry", "")
    print(f"✅ context 成功生成")
    print(f"   industry: {repr(industry)}")
    print(f"   industry 是否為空: {industry == ''}")
    print(f"   industry 是否為 None: {industry is None}")
    print(f"   industry 長度: {len(industry) if industry else 0}")
    print(f"   industry 類型: {type(industry)}")
    print(f"   所有 keys: {list(context.keys())}")
else:
    print("❌ context 為 None")

# Step 4: 測試實際使用
print("\n" + "="*80)
print("【Step 4】測試 PPTContentEngine 實際使用")
print("-"*80)
try:
    from content_pptx_company import PPTContentEngine
    engine = PPTContentEngine()
    
    industry = engine.env_context.get("industry", "")
    print(f"✅ PPTContentEngine 初始化成功")
    print(f"   env_context industry: {repr(industry)}")
    print(f"   industry 是否為空: {industry == ''}")
    print(f"   industry 是否為 None: {industry is None}")
    print(f"   industry 長度: {len(industry) if industry else 0}")
    print(f"   industry 類型: {type(industry)}")
    
    # 檢查是否會影響 prompt 生成
    if industry:
        intro = engine._format_expert_intro("本公司", industry)
        print(f"   生成的 intro: {repr(intro)}")
        if industry in intro:
            print(f"   ✅ 產業別已包含在 intro 中")
        else:
            print(f"   ❌ 產業別未包含在 intro 中")
    else:
        print(f"   ⚠️ 產業別為空，無法生成包含產業別的 intro")
        
except Exception as e:
    print(f"❌ 錯誤: {e}")
    import traceback
    traceback.print_exc()

# Step 5: 總結
print("\n" + "="*80)
print("【總結】")
print("-"*80)
if log_data and log_data.get("industry"):
    print("✅ log_data 中有產業別")
else:
    print("❌ log_data 中沒有產業別")
    
if context and context.get("industry"):
    print("✅ context 中有產業別")
else:
    print("❌ context 中沒有產業別")
    
try:
    from content_pptx_company import PPTContentEngine
    engine = PPTContentEngine()
    if engine.env_context.get("industry"):
        print("✅ PPTContentEngine.env_context 中有產業別")
    else:
        print("❌ PPTContentEngine.env_context 中沒有產業別")
except:
    pass

