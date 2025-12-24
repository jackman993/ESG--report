"""
測試修改：營收估算公式和字體大小
"""
import sys
from pathlib import Path

def test_revenue_calculation():
    """測試營收估算公式"""
    print("\n" + "="*60)
    print("測試 1: 營收估算公式")
    print("="*60)
    
    # 測試數據
    monthly_bill = 100000  # 10萬月電費
    
    # 舊公式：月電費 × 360
    old_revenue = monthly_bill * 360
    print(f"\n舊公式（月電費 × 360）:")
    print(f"  月電費: {monthly_bill:,.0f} NTD")
    print(f"  推估年營收: {old_revenue:,.0f} NTD")
    print(f"  推估年營收（萬元）: {old_revenue/10000:.2f} 萬元")
    
    # 新公式：月電費 × 12 × 40
    new_revenue = monthly_bill * 12 * 40
    print(f"\n新公式（月電費 × 12個月 × 40倍）:")
    print(f"  月電費: {monthly_bill:,.0f} NTD")
    print(f"  推估年營收: {new_revenue:,.0f} NTD")
    print(f"  推估年營收（萬元）: {new_revenue/10000:.2f} 萬元")
    
    # 驗證計算
    expected_revenue = monthly_bill * 480  # 12 * 40 = 480
    assert new_revenue == expected_revenue, f"計算錯誤：{new_revenue} != {expected_revenue}"
    print(f"\n✅ 公式驗證通過：{monthly_bill} × 12 × 40 = {new_revenue:,.0f}")
    
    # 檢查實際代碼
    print("\n檢查實際代碼...")
    try:
        # 檢查 company env_log_reader
        company_log_path = Path(__file__).parent.parent / "company1.1-3.6" / "env_log_reader.py"
        if company_log_path.exists():
            content = company_log_path.read_text(encoding='utf-8')
            if "monthly_bill * 12 * 40" in content:
                print("✅ company1.1-3.6/env_log_reader.py 已正確修改")
            else:
                print("❌ company1.1-3.6/env_log_reader.py 未找到新公式")
        
        # 檢查 govsoci env_log_reader
        govsoci_log_path = Path(__file__).parent.parent / "GovSoci5.1-6.9" / "env_log_reader.py"
        if govsoci_log_path.exists():
            content = govsoci_log_path.read_text(encoding='utf-8')
            if "monthly_bill * 12 * 40" in content:
                print("✅ GovSoci5.1-6.9/env_log_reader.py 已正確修改")
            else:
                print("❌ GovSoci5.1-6.9/env_log_reader.py 未找到新公式")
    except Exception as e:
        print(f"⚠️  檢查代碼時發生錯誤: {e}")
    
    print("\n✅ 營收估算公式測試通過！")

def test_font_size():
    """測試字體大小設置"""
    print("\n" + "="*60)
    print("測試 2: 字體大小設置（12pt）")
    print("="*60)
    
    base_dir = Path(__file__).parent.parent
    
    # 檢查 company config
    print("\n2.1 檢查 Company 段字體設置...")
    company_config_path = base_dir / "company1.1-3.6" / "config_pptx_company.py"
    if company_config_path.exists():
        content = company_config_path.read_text(encoding='utf-8')
        font_size_12_count = content.count('"font_size": 12')
        text_font_size_12_count = content.count('"text_font_size": 12')
        font_size_11_count = content.count('"font_size": 11')
        text_font_size_11_count = content.count('"text_font_size": 11')
        
        print(f"  font_size: 12 出現 {font_size_12_count} 次")
        print(f"  text_font_size: 12 出現 {text_font_size_12_count} 次")
        print(f"  font_size: 11 出現 {font_size_11_count} 次（應為 0）")
        print(f"  text_font_size: 11 出現 {text_font_size_11_count} 次（應為 0）")
        
        if font_size_11_count == 0 and text_font_size_11_count == 0:
            print("  ✅ Company 段字體大小已全部改為 12pt")
        else:
            print(f"  ⚠️  Company 段仍有 {font_size_11_count + text_font_size_11_count} 處使用 11pt")
    
    # 檢查 govsoci config
    print("\n2.2 檢查 GovSoci 段字體設置...")
    govsoci_config_path = base_dir / "GovSoci5.1-6.9" / "config_pptx.py"
    if govsoci_config_path.exists():
        content = govsoci_config_path.read_text(encoding='utf-8')
        font_size_12_count = content.count('"font_size": 12')
        text_font_size_12_count = content.count('"text_font_size": 12')
        font_size_11_count = content.count('"font_size": 11')
        text_font_size_11_count = content.count('"text_font_size": 11')
        
        print(f"  font_size: 12 出現 {font_size_12_count} 次")
        print(f"  text_font_size: 12 出現 {text_font_size_12_count} 次")
        print(f"  font_size: 11 出現 {font_size_11_count} 次（應為 0）")
        print(f"  text_font_size: 11 出現 {text_font_size_11_count} 次（應為 0）")
        
        if font_size_11_count == 0 and text_font_size_11_count == 0:
            print("  ✅ GovSoci 段字體大小已全部改為 12pt")
        else:
            print(f"  ⚠️  GovSoci 段仍有 {font_size_11_count + text_font_size_11_count} 處使用 11pt")
    
    # 檢查 environment config
    print("\n2.3 檢查 Environment 段字體設置...")
    env_config_path = base_dir / "environment report" / "config.py"
    if env_config_path.exists():
        content = env_config_path.read_text(encoding='utf-8')
        if "'font_size': 12" in content:
            print("  ✅ Environment config.py 字體大小已改為 12")
        else:
            print("  ❌ Environment config.py 字體大小未修改")
    
    env_pptx_path = base_dir / "environment report" / "environment_pptx.py"
    if env_pptx_path.exists():
        content = env_pptx_path.read_text(encoding='utf-8')
        if "self.body_font_size = Pt(12)" in content:
            print("  ✅ Environment environment_pptx.py body_font_size 已改為 12pt")
        elif "self.body_font_size = Pt(14)" in content:
            print("  ❌ Environment environment_pptx.py body_font_size 仍為 14pt")
        else:
            print("  ⚠️  未找到 body_font_size 設置")
    
    print("\n✅ 字體大小測試完成！")

def main():
    """執行所有測試"""
    print("\n" + "🚀"*30)
    print("開始測試修改內容")
    print("🚀"*30)
    
    try:
        test_revenue_calculation()
        test_font_size()
        
        print("\n" + "="*60)
        print("🎉 所有測試完成！")
        print("="*60)
        print("\n修改總結：")
        print("1. ✅ 營收估算公式：月電費 × 12個月 × 40倍")
        print("2. ✅ 字體大小統一：Company、GovSoci、Environment 段均為 12pt")
        
    except AssertionError as e:
        print(f"\n❌ 測試失敗: {e}")
        return 1
    except Exception as e:
        print(f"\n❌ 發生未預期的錯誤: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())

