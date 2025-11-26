import pandas as pd
import os
import time
import sys

# 强制使用UTF-8打印，防止控制台乱码
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("===========================================")
print("      智能对账系统 (云端打包版) v3.0")
print("===========================================")
print("正在初始化...")
time.sleep(1)

def read_csv_safe(filename):
    if not os.path.exists(filename):
        return None
    # 尝试多种编码读取
    encodings = ['gbk', 'utf-8', 'gb18030']
    for enc in encodings:
        try:
            return pd.read_csv(filename, encoding=enc)
        except:
            continue
    return None

def main():
    print(f"当前工作目录: {os.getcwd()}")
    print("请确保 '磅单列表.csv' 和 '门禁数据.csv' 与本程序在同一文件夹下。")
    
    df_bang = read_csv_safe('磅单列表.csv')
    df_men = read_csv_safe('门禁数据.csv')

    if df_bang is None or df_men is None:
        print("\n❌ 错误：未找到数据文件！")
        print("请检查文件名是否完全正确。")
        input("按回车键退出...")
        return

    print("✅ 文件读取成功，正在计算...")

    # 1. 清洗列名
    df_bang.columns = [c.strip() for c in df_bang.columns]
    df_men.columns = [c.strip() for c in df_men.columns]

    # 2. 关键列配置
    COL_BANG_PLATE = '车号'
    COL_BANG_NAME_D = '存货名称'
    COL_BANG_SPEC_E = '规格型号'
    COL_BANG_WEIGHT = '净重'
    COL_BANG_TIME = '毛重过磅时间'
    COL_MEN_PLATE = '车牌号'
    COL_MEN_NAME = '运输货物名称'
    COL_MEN_WEIGHT = '运输货物净重'
    COL_MEN_TIME = '出厂时间'

    # 3. 数据转换
    df_bang['过磅时间_dt'] = pd.to_datetime(df_bang[COL_BANG_TIME], errors='coerce')
    df_men['出厂时间_dt'] = pd.to_datetime(df_men[COL_MEN_TIME], errors='coerce')
    df_bang['净重_num'] = pd.to_numeric(df_bang[COL_BANG_WEIGHT], errors='coerce').fillna(0)
    df_men['门禁净重_num'] = pd.to_numeric(df_men[COL_MEN_WEIGHT], errors='coerce').fillna(0)
    
    # 去除车牌空格
    df_bang['join_plate'] = df_bang[COL_BANG_PLATE].astype(str).str.replace(' ', '').str.strip()
    df_men['join_plate'] = df_men[COL_MEN_PLATE].astype(str).str.replace(' ', '').str.strip()

    # 智能货名
    def get_real_name(row):
        d = str(row.get(COL_BANG_NAME_D, '')).strip()
        e = str(row.get(COL_BANG_SPEC_E, '')).strip()
        if not d or d.lower() == 'nan' or d.replace('.','').isdigit():
            return e
        return d
    df_bang['有效货名'] = df_bang.apply(get_real_name, axis=1)

    results = []

    for index, row in df_bang.iterrows():
        plate = row['join_plate']
        p_time = row['过磅时间_dt']
        p_weight = row['净重_num']
        
        res = {
            '门禁车牌': '未找到',
            '门禁出厂时间': '',
            '门禁净重': '',
            '门禁货名': '',
            '状态_车牌': '🔴异常',
            '状态_净重': '🔴异常',
            '状态_货名': '🔴异常',
            '备注': ''
        }

        if pd.isnull(p_time):
            res['备注'] = '磅单时间无效'
        else:
            # 查找车牌
            subset = df_men[df_men['join_plate'] == plate].copy()
            if not subset.empty:
                # 查找时间：出厂时间 >= 过磅时间
                future = subset[subset['出厂时间_dt'] >= p_time].copy()
                if not future.empty:
                    future['diff'] = future['出厂时间_dt'] - p_time
                    best = future.sort_values('diff').iloc[0]
                    
                    diff_min = best['diff'].total_seconds() / 60
                    if diff_min < 2880: # 48小时内
                        res['门禁车牌'] = best[COL_MEN_PLATE]
                        res['门禁出厂时间'] = best[COL_MEN_TIME]
                        res['门禁净重'] = best['门禁净重_num']
                        res['门禁货名'] = best[COL_MEN_NAME]
                        
                        res['状态_车牌'] = '🟢正常'
                        
                        # 比对重量
                        if abs(p_weight - best['门禁净重_num']) <= 0.02:
                            res['状态_净重'] = '🟢正常'
                        else:
                            res['状态_净重'] = '🟡不符'
                            
                        # 比对货名
                        m_name = str(best[COL_MEN_NAME])
                        p_name = str(row['有效货名'])
                        if m_name == p_name or m_name in p_name or p_name in m_name:
                            res['状态_货名'] = '🟢正常'
                        else:
                            # 模糊匹配
                            kws = ['焦','煤','油','酸','碱','盐','苯']
                            is_fuzzy = False
                            for kw in kws:
                                if kw in m_name and kw in p_name:
                                    is_fuzzy = True
                                    break
                            res['状态_货名'] = '🟢模糊匹配' if is_fuzzy else '🔴不符'
                    else:
                        res['备注'] = '时间间隔过长(>48h)'
                else:
                    res['备注'] = '无后续出厂记录'
        
        row_data = row.to_dict()
        # 清理临时数据
        for k in ['过磅时间_dt', '净重_num', 'join_plate', '有效货名']:
            if k in row_data: del row_data[k]
        row_data.update(res)
        results.append(row_data)

    # 保存
    output_name = '最终对账结果.csv'
    try:
        pd.DataFrame(results).to_csv(output_name, index=False, encoding='gbk') # Excel用GBK
        print(f"\n✅ 成功！结果已保存为: {output_name}")
    except:
        print("\n⚠️ 保存失败，文件可能被占用")

    print("\n程序运行完毕，请查看生成的CSV文件。")
    input("按回车键退出...")

if __name__ == '__main__':
    main()
