import pandas as pd
import os
import time
import sys
import io

# ---------------------------------------------------------
# 基础设置：防止中文乱码
# ---------------------------------------------------------
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
except:
    pass

print("===========================================")
print("      智能对账系统 V6.0 (终极稳定版)")
print("===========================================")
print(f"当前工作目录: {os.getcwd()}\n")

# ---------------------------------------------------------
# 核心工具函数
# ---------------------------------------------------------
def clean_text(text):
    """强力清洗：去除BOM头、幽灵空格、普通空格"""
    if pd.isna(text):
        return ""
    s = str(text)
    # 去除 BOM (\ufeff), 不间断空格 (\xa0), 全角空格 (\u3000) 和普通空格
    s = s.replace('\ufeff', '').replace('\xa0', '').replace('\u3000', '').replace(' ', '')
    return s

def read_csv_smart(filename):
    """尝试多种编码读取文件"""
    if not os.path.exists(filename):
        return None
    encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb18030']
    for enc in encodings:
        try:
            df = pd.read_csv(filename, encoding=enc)
            # 读取成功后，立刻清洗列名，防止 BOM 头导致列名无法识别
            df.columns = [clean_text(c) for c in df.columns]
            return df
        except:
            continue
    return None

def main():
    print("1. 正在读取文件...")
    
    df_bang = read_csv_smart('磅单列表.csv')
    df_men = read_csv_smart('门禁数据.csv')

    if df_bang is None or df_men is None:
        print("❌ 错误：未找到 '磅单列表.csv' 或 '门禁数据.csv'")
        input("按回车键退出...")
        return

    # ---------------------------------------------------------
    # 关键配置：对应你提供的数据列名
    # ---------------------------------------------------------
    # 磅单表列名
    K_BANG_PLATE = '车号'
    K_BANG_NAME_D = '存货名称'
    K_BANG_SPEC_E = '规格型号'
    K_BANG_WEIGHT = '净重'
    K_BANG_TIME = '毛重过磅时间'
    
    # 门禁表列名
    K_MEN_PLATE = '车牌号'
    K_MEN_NAME = '运输货物名称'
    K_MEN_WEIGHT = '运输货物净重'
    K_MEN_TIME = '出厂时间'

    # 检查列是否存在
    missing_cols = []
    for c in [K_BANG_PLATE, K_BANG_WEIGHT, K_BANG_TIME]:
        if c not in df_bang.columns: missing_cols.append(f"磅单列表缺[{c}]")
    for c in [K_MEN_PLATE, K_MEN_WEIGHT, K_MEN_TIME]:
        if c not in df_men.columns: missing_cols.append(f"门禁数据缺[{c}]")
    
    if missing_cols:
        print("❌ 列名不匹配错误：")
        for m in missing_cols: print(f"   - {m}")
        print("请检查CSV文件的表头是否正确。")
        input("按回车键退出...")
        return

    print(f"✅ 读取成功！磅单记录: {len(df_bang)} 条 | 门禁记录: {len(df_men)} 条")
    print("2. 正在进行全表清洗 (去除隐形字符)...")

    # 统一全表清洗函数
    def clean_df_content(df):
        # 只清洗字符串列
        obj_cols = df.select_dtypes(include=['object']).columns
        for col in obj_cols:
            df[col] = df[col].apply(lambda x: str(x).replace('\xa0', ' ').strip() if pd.notnull(x) else x)
        return df

    df_bang = clean_df_content(df_bang)
    df_men = clean_df_content(df_men)

    # 转换关键数据类型
    try:
        df_bang['dt_base'] = pd.to_datetime(df_bang[K_BANG_TIME], errors='coerce')
        df_men['dt_base'] = pd.to_datetime(df_men[K_MEN_TIME], errors='coerce')
        
        df_bang['num_weight'] = pd.to_numeric(df_bang[K_BANG_WEIGHT], errors='coerce').fillna(0)
        df_men['num_weight'] = pd.to_numeric(df_men[K_MEN_WEIGHT], errors='coerce').fillna(0)
        
        # 用于匹配的车牌：去空格、大写
        df_bang['key_plate'] = df_bang[K_BANG_PLATE].astype(str).str.replace(' ', '').str.upper()
        df_men['key_plate'] = df_men[K_MEN_PLATE].astype(str).str.replace(' ', '').str.upper()
    except Exception as e:
        print(f"❌ 数据格式转换错误: {e}")
        input("按回车键退出...")
        return

    # 智能货名提取
    def get_real_name(row):
        d = str(row.get(K_BANG_NAME_D, '')).strip()
        e = str(row.get(K_BANG_SPEC_E, '')).strip()
        # 如果D是数字或空，取E
        if not d or d.lower() == 'nan' or d.replace('.', '').isdigit():
            return e
        return d
    
    df_bang['有效货名'] = df_bang.apply(get_real_name, axis=1)

    # ---------------------------------------------------------
    # 核心比对循环
    # ---------------------------------------------------------
    print("3. 开始智能匹配 (逻辑：匹配车牌 -> 找过磅后的最近出厂时间)...")
    
    results = []
    
    # 预先过滤有效门禁数据，提升速度
    valid_men = df_men.dropna(subset=['dt_base'])

    for index, row in df_bang.iterrows():
        plate = row['key_plate']
        time_bang = row['dt_base']
        weight_bang = row['num_weight']
        name_bang = row['有效货名']
        
        res = {
            '门禁_匹配车牌': '未找到',
            '门禁_出厂时间': '',
            '门禁_货物名称': '',
            '门禁_净重': 0,
            '停留时长(分)': '',
            '结果_车牌': '🔴异常',
            '结果_净重': '🔴异常',
            '结果_货名': '🔴异常',
            '备注': ''
        }

        if pd.isnull(time_bang):
            res['备注'] = '磅单时间格式错误'
        else:
            # 1. 筛选车牌
            matches = valid_men[valid_men['key_plate'] == plate]
            
            if not matches.empty:
                # 2. 筛选时间：出厂时间 >= 过磅时间
                future_matches = matches[matches['dt_base'] >= time_bang].copy()
                
                if not future_matches.empty:
                    # 3. 计算时间差，取最小的
                    future_matches['diff'] = future_matches['dt_base'] - time_bang
                    best = future_matches.sort_values('diff').iloc[0]
                    
                    diff_minutes = best['diff'].total_seconds() / 60
                    
                    # 填充基础信息
                    res['门禁_匹配车牌'] = best[K_MEN_PLATE]
                    res['门禁_出厂时间'] = best[K_MEN_TIME]
                    res['门禁_货物名称'] = best[K_MEN_NAME]
                    res['门禁_净重'] = best['num_weight']
                    res['停留时长(分)'] = round(diff_minutes, 1)
                    
                    # --- 判定逻辑 ---
                    
                    # A. 车牌判定 (能进到这里说明肯定一致)
                    if diff_minutes > 2880: # 48小时
                        res['结果_车牌'] = '🟡时间过长'
                        res['备注'] = '停留超过48小时'
                    else:
                        res['结果_车牌'] = '🟢正常'
                        
                    # B. 净重判定 (误差0.02)
                    if abs(weight_bang - best['num_weight']) <= 0.02:
                        res['结果_净重'] = '🟢正常'
                    else:
                        res['结果_净重'] = '🟡不符'
                        
                    # C. 货名判定 (模糊匹配)
                    m_name = str(best[K_MEN_NAME])
                    p_name = str(name_bang)
                    
                    # 关键词库
                    keywords = ['焦', '煤', '油', '酸', '碱', '盐', '苯']
                    is_fuzzy_match = False
                    
                    if m_name == p_name or m_name in p_name or p_name in m_name:
                        is_fuzzy_match = True
                    else:
                        for kw in keywords:
                            if kw in m_name and kw in p_name:
                                is_fuzzy_match = True
                                break
                    
                    if is_fuzzy_match:
                        res['结果_货名'] = '🟢正常'
                    else:
                        res['结果_货名'] = '🔴不符'
                        res['备注'] += f" | 名:磅[{p_name}]/门[{m_name}]"
                        
                else:
                    res['备注'] = '无过磅后的出厂记录'
            else:
                res['备注'] = '门禁无此车牌'

        # 合并数据
        row_data = row.to_dict()
        # 清理过程数据
        for k in ['dt_base', 'num_weight', 'key_plate', 'joine_plate', '有效货名']:
             if k in row_data: del row_data[k]
        
        row_data.update(res)
        results.append(row_data)

    # ---------------------------------------------------------
    # 导出保存 (防冲突机制)
    # ---------------------------------------------------------
    # 调整列顺序
    df_final = pd.DataFrame(results)
    
    # 把结果列提前，方便查看
    cols_order = [K_BANG_PLATE, '结果_车牌', '结果_净重', '结果_货名', '停留时长(分)', '备注', 
                  '门禁_匹配车牌', '门禁_净重', '门禁_货物名称', '门禁_出厂时间']
    # 加上原表其他列
    remaining_cols = [c for c in df_final.columns if c not in cols_order]
    df_final = df_final[cols_order + remaining_cols]

    # 生成带时间戳的文件名，避免“文件被占用”错误
    timestamp = time.strftime("%H点%M分%S秒")
    output_filename = f'对账结果_{timestamp}.csv'
    
    print(f"4. 正在保存为: {output_filename} ...")
    
    try:
        # 使用 utf-8-sig 编码，彻底解决乱码和保存崩溃问题
        df_final.to_csv(output_filename, index=False, encoding='utf-8-sig')
        print(f"\n✅✅✅ 全部完成！请打开 [{output_filename}] 查看结果。")
    except Exception as e:
        print(f"\n❌ 保存失败: {e}")

    print("\n(请按回车键关闭此窗口)")
    input()

if __name__ == '__main__':
    main()
