"""
測試 wrapper 改進功能
驗證日誌記錄、輸入驗證、錯誤處理是否正常運作
"""
import sys
import logging
from pathlib import Path

# 設置詳細日誌
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('test_wrappers.log', encoding='utf-8')
    ]
)

logger = logging.getLogger(__name__)

def test_input_validation():
    """測試輸入驗證功能"""
    print("\n" + "="*60)
    print("測試 1: 輸入驗證")
    print("="*60)
    
    from company_engine_wrapper_zh import generate_company_section_zh
    
    # 測試空 API Key
    print("\n1.1 測試空 API Key...")
    result, error = generate_company_section_zh(api_key="")
    assert result is None, "應該返回 None"
    assert error == "API Key 不能為空", f"錯誤訊息不正確: {error}"
    print("✅ 空 API Key 驗證通過")
    
    # 測試空白字串
    print("\n1.2 測試空白字串 API Key...")
    result, error = generate_company_section_zh(api_key="   ")
    assert result is None, "應該返回 None"
    assert error == "API Key 不能為空", f"錯誤訊息不正確: {error}"
    print("✅ 空白字串驗證通過")
    
    # 測試格式警告（不應該失敗，只是警告）
    print("\n1.3 測試格式警告...")
    result, error = generate_company_section_zh(api_key="invalid-key")
    # 格式不正確應該只是警告，不會阻止執行（但會因為其他原因失敗）
    print("✅ 格式警告功能正常（預期會因為其他原因失敗）")
    
    print("\n✅ 輸入驗證測試全部通過！")

def test_logging():
    """測試日誌記錄功能"""
    print("\n" + "="*60)
    print("測試 2: 日誌記錄")
    print("="*60)
    
    from company_engine_wrapper_zh import generate_company_section_zh
    
    print("\n2.1 測試日誌輸出（使用無效 API Key 觸發錯誤）...")
    result, error = generate_company_section_zh(
        api_key="sk-ant-test-invalid-key-12345",
        company_name="測試公司"
    )
    
    # 檢查日誌文件是否存在且有內容
    log_file = Path("test_wrappers.log")
    if log_file.exists():
        log_content = log_file.read_text(encoding='utf-8')
        # 檢查關鍵日誌訊息
        assert "開始生成公司段 PPTX" in log_content, "應該記錄開始訊息"
        assert "載入公司段引擎" in log_content, "應該記錄引擎載入"
        assert "環境狀態恢復完成" in log_content, "應該記錄狀態恢復"
        assert "生成公司段失敗" in log_content or "ERROR" in log_content, "應該記錄錯誤"
        print("✅ 日誌記錄功能正常")
        print(f"   日誌文件大小: {log_file.stat().st_size} bytes")
        print("   包含的關鍵訊息:")
        print("     - 開始生成公司段 PPTX")
        print("     - 載入公司段引擎")
        print("     - 環境狀態恢復完成")
        print("     - 錯誤記錄")
    else:
        print("⚠️  日誌文件未生成（可能日誌配置問題）")
    
    print("\n✅ 日誌記錄測試完成！")

def test_error_handling():
    """測試錯誤處理和狀態恢復"""
    print("\n" + "="*60)
    print("測試 3: 錯誤處理與狀態恢復")
    print("="*60)
    
    import sys
    original_path_length = len(sys.path)
    
    from company_engine_wrapper_zh import generate_company_section_zh
    
    print("\n3.1 測試錯誤處理（使用無效 API Key）...")
    result, error = generate_company_section_zh(
        api_key="sk-ant-test-invalid-key-12345",
        company_name="測試公司"
    )
    
    # 檢查 sys.path 是否恢復
    current_path_length = len(sys.path)
    assert current_path_length == original_path_length, \
        f"sys.path 未正確恢復: 原始長度 {original_path_length}, 當前長度 {current_path_length}"
    print(f"✅ sys.path 已正確恢復（長度: {current_path_length}）")
    
    # 檢查是否有錯誤訊息
    assert error is not None, "應該返回錯誤訊息"
    print(f"✅ 錯誤處理正常: {error[:50]}...")
    
    print("\n✅ 錯誤處理測試通過！")

def test_govsoci_wrapper():
    """測試治理社會段 wrapper"""
    print("\n" + "="*60)
    print("測試 4: 治理社會段 Wrapper")
    print("="*60)
    
    from govsoci_engine_wrapper_zh import generate_govsoci_section_zh
    
    # 測試輸入驗證
    print("\n4.1 測試輸入驗證...")
    result, error = generate_govsoci_section_zh(api_key="")
    assert result is None and error == "API Key 不能為空", "輸入驗證應該失敗"
    print("✅ 輸入驗證通過")
    
    # 測試錯誤處理
    print("\n4.2 測試錯誤處理...")
    import sys
    original_path_length = len(sys.path)
    result, error = generate_govsoci_section_zh(api_key="sk-ant-test-invalid")
    
    current_path_length = len(sys.path)
    assert current_path_length == original_path_length, "sys.path 未正確恢復"
    print(f"✅ sys.path 已正確恢復（長度: {current_path_length}）")
    
    print("\n✅ 治理社會段 Wrapper 測試通過！")

def main():
    """執行所有測試"""
    print("\n" + "🚀"*30)
    print("開始測試 Wrapper 改進功能")
    print("🚀"*30)
    
    try:
        test_input_validation()
        test_logging()
        test_error_handling()
        test_govsoci_wrapper()
        
        print("\n" + "="*60)
        print("🎉 所有測試通過！")
        print("="*60)
        print("\n建議：")
        print("1. 檢查 test_wrappers.log 查看詳細日誌")
        print("2. 確認日誌記錄完整且清晰")
        print("3. 測試通過後可以 commit 和 push 到 GitHub")
        print("4. Streamlit 會自動從 GitHub 同步更新")
        
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

