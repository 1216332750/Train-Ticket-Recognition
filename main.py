import fitz  # PyMuPDF
import os
import re
import shutil
import pandas as pd
from rapidocr_onnxruntime import RapidOCR

# ================= 配置区 =================
INPUT_DIR = './pdfs'           # 原始 PDF 文件夹路径
RENAMED_DIR = './重命名后发票' # 重命名后的 PDF 存放文件夹
OUTPUT_FILE = '铁路发票精准提取结果.xlsx'
# ==========================================

engine = RapidOCR()

def get_pdf_content_by_file(root_dir):
    """提取纯净文本块并记录原始文件路径"""
    all_files_results = []
    if not os.path.exists(root_dir):
        print(f"❌ 目录 {root_dir} 不存在，请检查。")
        return []

    print(f"🔍 正在扫描 {root_dir} 下的 PDF 文件...\n")
    for root, dirs, files in os.walk(root_dir):
        for file in sorted(files):
            if file.lower().endswith('.pdf'):
                file_path = os.path.join(root, file)
                current_file_data = []
                try:
                    doc = fitz.open(file_path)
                    for page in doc:
                        blocks = page.get_text("blocks")
                        if not blocks or (len(blocks) == 1 and not blocks[0][4].strip()):
                            pix = page.get_pixmap(dpi=150)
                            result, _ = engine(pix.tobytes())
                            if result:
                                for line in result:
                                    clean_text = re.sub(r'\s+', '', line[1])
                                    if clean_text: current_file_data.append(clean_text)
                        else:
                            blocks.sort(key=lambda b: (b[1], b[0]))
                            for b in blocks:
                                clean_text = re.sub(r'\s+', '', b[4])
                                if clean_text: current_file_data.append(clean_text)
                    doc.close()
                    all_files_results.append({
                        "filename": file,
                        "filepath": file_path, 
                        "data": current_file_data
                    })
                except Exception as e:
                    print(f"⚠️ 处理文件 {file} 出错: {e}")
    return all_files_results

def parse_railway_data_optimal(all_files_results):
    """精准抓取业务字段"""
    final_table = []
    print("🧠 正在分析内容并提取结构化字段...")

    for file_item in all_files_results:
        filename = file_item['filename']
        filepath = file_item['filepath']
        blocks = file_item['data']
        
        valid_blocks = [b for b in blocks if len(b) > 1]
        original_text = "".join(valid_blocks)
        
        amount_match = re.search(r'￥(\d+\.\d{2})', original_text)
        amount = amount_match.group(1) if amount_match else ""

        safe_text = re.sub(r'￥\d+\.\d{2}', ' ', original_text)
        invoice_type = "退票发票" if "退票" in safe_text else "购票发票"

        person_match = re.search(r'([A-Za-z0-9]+\*{3,}(?:\d{1,4}[Xx]?)?)([A-Za-z\u4e00-\u9fa5]+)', safe_text)
        if person_match:
            id_card = person_match.group(1)
            raw_name = person_match.group(2)
            name = raw_name.replace("电子客票号", "").replace("退票费", "").replace("退票", "")
        else:
            id_card, name = "", ""

        station_train_match = re.search(r'([\u4e00-\u9fa5]+站)[A-Za-z]*([\u4e00-\u9fa5]+站)([A-Z]\d{1,4})?', safe_text)
        if station_train_match:
            depart_station = station_train_match.group(1)
            arrive_station = station_train_match.group(2)
            train_no = station_train_match.group(3) if station_train_match.group(3) else ""
        else:
            stations = re.findall(r'([\u4e00-\u9fa5]+站)', safe_text)
            depart_station = stations[0] if len(stations) > 0 else ""
            arrive_station = stations[1] if len(stations) > 1 else ""
            train_no = ""

        if not train_no:
            train_no_match = re.search(r'([A-Z]\d{1,4})', safe_text)
            train_no = train_no_match.group(1) if train_no_match else ""

        time_seat_match = re.search(r'(\d{4}年\d{2}月\d{2}日)(\d{2}:\d{2})开(\d+车[A-Z0-9]+号)', original_text)
        if time_seat_match:
            date, time, seat_no = time_seat_match.group(1), time_seat_match.group(2), time_seat_match.group(3)
        else:
            date_match = re.search(r'(\d{4}年\d{2}月\d{2}日)', original_text)
            date = date_match.group(1) if date_match else ""
            time, seat_no = "", ""

        invoice_no_match = re.search(r'发票号码:(\d{20})', original_text)
        invoice_no = invoice_no_match.group(1) if invoice_no_match else ""
        
        invoice_date_match = re.search(r'开票日期:(\d{4}年\d{2}月\d{2}日)', original_text)
        invoice_date = invoice_date_match.group(1) if invoice_date_match else ""

        seat_class_match = re.search(r'(商务座|特等座|一等座|二等座|无座)', original_text)
        seat_class = seat_class_match.group(1) if seat_class_match else ""

        ticket_no_match = re.search(r'电子客票号:(\d+)', safe_text)
        ticket_no = ticket_no_match.group(1) if ticket_no_match else ""

        buyer_match = re.search(r'购买方名称:([\u4e00-\u9fa5]+)统一社会信用代码:([0-9A-Z]{18})', safe_text)
        buyer_name = buyer_match.group(1) if buyer_match else ""
        buyer_tax_id = buyer_match.group(2) if buyer_match else ""

        final_table.append({
            "原路径": filepath, 
            "来源文件": filename,
            "发票类型": invoice_type,
            "发票号码": invoice_no,
            "开票日期": invoice_date,
            "乘车人": name,
            "身份证号": id_card,
            "出发站": depart_station,
            "到达站": arrive_station,
            "车次": train_no,
            "乘车日期": date,
            "发车时间": time,
            "坐席等级": seat_class,
            "车厢及座位": seat_no,
            "金额(元)": amount,
            "电子客票号": ticket_no,
            "购买方公司": buyer_name,
            "税号": buyer_tax_id
        })
    return final_table

# ================= 主程序入口 =================
if __name__ == "__main__":
    all_files_results = get_pdf_content_by_file(INPUT_DIR)
    
    if all_files_results:
        structured_data = parse_railway_data_optimal(all_files_results)
        df = pd.DataFrame(structured_data)
        
        # ---------------------------------------------------------
        # 核心逻辑区：严格落实“先排序 -> 生成新序号 -> 重命名 -> 输出Excel”
        # ---------------------------------------------------------

        # 【动作1：先按同名排序】按乘车人A-Z排序，同一个人的票再按日期排序
        print("\n🗂️ 正在按乘车人同名归类排序...")
        df = df.sort_values(by=["乘车人", "乘车日期"], ascending=[True, True]).reset_index(drop=True)
        
        # 【动作2：打上新序号】因为已经排好序了，现在的第1行就是真正的第1个文件
        df.insert(0, '序号', df.index + 1)
        
        # 【动作3：重命名文件输出】按照这个带有新序号的顺序，挨个给PDF重命名
        print("📁 正在生成新序号并重命名文件...")
        os.makedirs(RENAMED_DIR, exist_ok=True)
        
        for index, row in df.iterrows():
            old_path = row["原路径"]
            xuhao = row["序号"]
            passenger = row["乘车人"] if row["乘车人"] else "未知姓名"
            amount = row["金额(元)"] if row["金额(元)"] else "0.00"
            inv_type = row["发票类型"]
            
            # 文件名格式：序号-乘车人-金额-发票类型.pdf
            new_filename = f"{xuhao}-{passenger}-{amount}-{inv_type}.pdf"
            new_path = os.path.join(RENAMED_DIR, new_filename)
            
            try:
                shutil.copy2(old_path, new_path)
            except Exception as e:
                print(f"⚠️ 复制并重命名文件 {old_path} 失败: {e}")

        # 【动作4：输出Excel】清理掉程序用的原路径字段，把这张完美的表存下来
        df.drop(columns=["原路径"], inplace=True)
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"✅ 文件重命名完成！全部保存在 [{RENAMED_DIR}] 文件夹中。")
        print(f"🎉 最终排序Excel已导出至：{OUTPUT_FILE}\n")
