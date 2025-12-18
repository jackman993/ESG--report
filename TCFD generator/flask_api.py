#!/usr/bin/env python3
"""
Flask API for Logit Analysis
啟動方式: python app.py
訪問: http://localhost:5000
"""

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import io
import json
from generic_mixed_logit import GenericMixedLogit

app = Flask(__name__)
CORS(app)  # 允許跨域請求

@app.route('/')
def index():
    """首頁"""
    return send_from_directory('.', 'index.html')

@app.route('/api/logit', methods=['POST'])
def run_logit():
    """
    執行 Logit 分析
    
    請求格式:
    {
        "csv_data": "csv字串或base64",
        "model_type": "logit" 或 "mixed_logit",
        "id_col": "respondent_id",
        "alt_col": "alternative",
        "choice_col": "choice",
        "X_cols": ["price", "screen_size", ...],
        "random_params": ["price"] (可選，僅用於 mixed_logit)
    }
    """
    try:
        data = request.json
        
        # 解析 CSV 資料
        csv_data = data.get('csv_data')
        if not csv_data:
            return jsonify({'error': '缺少 CSV 資料'}), 400
        
        df = pd.read_csv(io.StringIO(csv_data))
        
        # 獲取參數
        model_type = data.get('model_type', 'logit')
        id_col = data.get('id_col')
        alt_col = data.get('alt_col')
        choice_col = data.get('choice_col')
        X_cols = data.get('X_cols')
        random_params = data.get('random_params', None)
        n_draws = data.get('n_draws', 500)
        
        # 驗證參數
        if not all([id_col, alt_col, choice_col, X_cols]):
            return jsonify({'error': '缺少必要參數'}), 400
        
        # 建立模型
        model = GenericMixedLogit(n_draws=n_draws, random_params=random_params)
        
        # 擬合模型
        results = model.fit(
            data=df,
            id_col=id_col,
            alt_col=alt_col,
            choice_col=choice_col,
            X_cols=X_cols,
            model_type=model_type
        )
        
        # 格式化結果
        response = format_results(results)
        
        return jsonify(response)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/example-data', methods=['GET'])
def get_example_data():
    """獲取範例資料"""
    try:
        # 讀取範例 CSV
        df = pd.read_csv('smartphone_choice.csv')
        
        return jsonify({
            'csv_data': df.to_csv(index=False),
            'metadata': {
                'n_rows': len(df),
                'n_respondents': df['respondent_id'].nunique(),
                'alternatives': df['alternative'].unique().tolist(),
                'variables': df.columns.tolist()
            }
        })
    except FileNotFoundError:
        return jsonify({'error': '找不到範例資料檔案'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def format_results(results):
    """格式化結果供前端使用"""
    model_type = results.get('model_type')
    
    if model_type == 'logit':
        mnl = results['mnl_result']
        
        # 係數表
        coefficients = []
        for i, var in enumerate(results['X_cols']):
            coefficients.append({
                'variable': var,
                'coefficient': float(mnl['beta'][i]),
                'std_error': None,  # 需要 Hessian 矩陣計算
                'p_value': None
            })
        
        # ASC
        for i, alt in enumerate(results['alternatives']):
            if i == 0:
                continue  # 基準
            coefficients.append({
                'variable': f'ASC_{alt}',
                'coefficient': float(mnl['asc'][i]),
                'std_error': None,
                'p_value': None
            })
        
        return {
            'model_type': 'logit',
            'log_likelihood': float(mnl['log_likelihood']),
            'accuracy': float(mnl['accuracy']),
            'coefficients': coefficients,
            'alternatives': results['alternatives'],
            'success': True
        }
    
    elif model_type == 'mixed_logit':
        mnl = results['mnl_result']
        mixed = results['mixed_result']
        
        if not mixed.get('success'):
            return {
                'model_type': 'mixed_logit',
                'success': False,
                'error': 'Mixed Logit 估計失敗'
            }
        
        # 固定係數
        fixed_coefs = []
        param_idx = 0
        X_cols = results['X_cols']
        
        for idx in mixed['fixed_idx']:
            fixed_coefs.append({
                'variable': X_cols[idx],
                'coefficient': float(mixed['params'][param_idx]),
                'type': 'fixed'
            })
            param_idx += 1
        
        # 隨機係數
        random_coefs = []
        n_random = mixed['n_random']
        
        for i, idx in enumerate(mixed['random_idx']):
            mean = float(mixed['params'][param_idx])
            std = float(mixed['params'][param_idx + n_random])
            random_coefs.append({
                'variable': X_cols[idx],
                'mean': mean,
                'std': std,
                'type': 'random'
            })
            param_idx += 1
        param_idx += n_random
        
        # ASC
        asc_coefs = []
        alternatives = results['alternatives']
        asc_coefs.append({
            'alternative': alternatives[0],
            'coefficient': 0.0,
            'note': '基準'
        })
        
        for i in range(1, len(alternatives)):
            asc_coefs.append({
                'alternative': alternatives[i],
                'coefficient': float(mixed['params'][param_idx])
            })
            param_idx += 1
        
        return {
            'model_type': 'mixed_logit',
            'log_likelihood': float(mixed['log_likelihood']),
            'accuracy': float(mixed['accuracy']),
            'baseline_accuracy': float(mnl['accuracy']),
            'improvement': float(mixed['accuracy'] - mnl['accuracy']),
            'fixed_coefficients': fixed_coefs,
            'random_coefficients': random_coefs,
            'asc': asc_coefs,
            'alternatives': alternatives,
            'success': True
        }

if __name__ == '__main__':
    print("=== Logit Analysis API ===")
    print("啟動 Flask 服務於 http://localhost:5000")
    print("\nAPI 端點:")
    print("  POST /api/logit - 執行 Logit 分析")
    print("  GET  /api/example-data - 獲取範例資料")
    print("\n按 Ctrl+C 停止服務")
    
    app.run(debug=True, host='0.0.0.0', port=5000)