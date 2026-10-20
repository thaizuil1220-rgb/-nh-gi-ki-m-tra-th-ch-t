import cv2
import mediapipe as mp
import numpy as np
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
from collections import deque
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
import time
from datetime import datetime

# ================= 1. CẤU HÌNH BÀI TẬP & TIÊU CHUẨN KỸ THUẬT =================
EXERCISES = {
    "Squat / Bật nhảy bổ trợ": {
        "description": "Kiểm tra độ sâu khớp gối và tư thế thẳng lưng",
        "prep_angle": 160,
        "action_min": 80,
        "action_max": 105,
        "back_limit": 75
    },
    "Xuất phát thấp (Tư thế 'Sẵn sàng')": {
        "description": "Gối trước ~90°, gối sau 110°-130°, mông cao hơn vai",
        "front_knee_min": 85,
        "front_knee_max": 100
    },
    "Cầu lông - Đập cầu thuận tay (Smash)": {
        "description": "Kiểm tra nâng khuỷu tay (nách >= 90°) và duỗi thẳng tay (>= 165°) khi đập cầu",
        "prep_shoulder_min": 90,
        "hit_elbow_min": 165
    },
    "Bóng rổ - Ném rổ 1 tay trên vai": {
        "description": "Kiểm tra góc khuỷu tay 90° khi nạp lực và duỗi thẳng tay (>= 165°) khi ra bóng",
        "dip_elbow_min": 85,
        "dip_elbow_max": 95,
        "release_elbow_min": 165
    }
}

OUTPUT_DIR = "hinh_anh_dong_tac_loi"
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# ================= 2. BỘ LỌC LÀM MỊN & TÍNH GÓC =================
class AngleSmoother:
    def __init__(self, window_size=5):
        self.history = deque(maxlen=window_size)
    def update(self, new_val):
        self.history.append(new_val)
        return round(sum(self.history) / len(self.history), 1)
    def reset(self):
        self.history.clear()

def calculate_angle(a, b, c):
    a, b, c = np.array(a), np.array(b), np.array(c)
    radians = np.arctan2(c[1] - b[1], c[0] - b[0]) - np.arctan2(a[1] - b[1], a[0] - b[0])
    angle = np.abs(radians * 180.0 / np.pi)
    if angle > 180.0:
        angle = 360.0 - angle
    return angle

def export_excel_with_chart(session_records, output_filepath):
    if not session_records:
        return False, "Không có dữ liệu để xuất!"

    wb = openpyxl.Workbook()
    ws_detail = wb.active
    ws_detail.title = "Nhat_Ky_Chi_Tiet"
    ws_detail.views.sheetView[0].showGridLines = True

    headers = ["Thời gian", "Họ tên", "Lớp", "Nội dung bài tập", "Lần thứ", "Kết quả", "Chi tiết lỗi", "Góc đo được", "Tiêu chuẩn"]
    ws_detail.append(headers)

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    for col_num in range(1, len(headers) + 1):
        cell = ws_detail.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    fill_pass = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_pass = Font(color="006100", bold=True)
    fill_fail = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_fail = Font(color="9C0006", bold=True)

    error_types_set = set()
    for row_idx, r in enumerate(session_records, start=2):
        ws_detail.append([
            r["Thời gian"], r["Họ tên"], r["Lớp"], r["Nội dung bài tập"],
            r["Lần thứ"], r["Kết quả"], r["Chi tiết lỗi"], r["Góc đo được"], r["Tiêu chuẩn"]
        ])
        if r["Kết quả"] == "LỖI" and r["Chi tiết lỗi"]:
            error_types_set.add(r["Chi tiết lỗi"])

        res_cell = ws_detail.cell(row=row_idx, column=6)
        if res_cell.value == "ĐẠT":
            res_cell.fill = fill_pass
            res_cell.font = font_pass
        else:
            res_cell.fill = fill_fail
            res_cell.font = font_fail

        for col_num in range(1, len(headers) + 1):
            ws_detail.cell(row=row_idx, column=col_num).border = thin_border

    for col in ws_detail.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_detail.column_dimensions[col_letter].width = max(max_len + 3, 12)

    total_rows = len(session_records)
    ws_dash = wb.create_sheet(title="Bao_Cao_Bieu_Do")
    ws_dash.views.sheetView[0].showGridLines = True

    ws_dash.merge_cells("B2:J2")
    title_cell = ws_dash["B2"]
    title_cell.value = "BÁO CÁO ĐÁNH GIÁ KỸ THUẬT & PHÂN TÍCH LỖI SAI THỂ DỤC THỂ THAO"
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="1F497D")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws_dash["B4"], ws_dash["C4"], ws_dash["D4"] = "Kết quả tổng quát", "Số lượt", "Tỷ lệ"
    for col in ["B4", "C4", "D4"]:
        ws_dash[col].fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        ws_dash[col].font = Font(bold=True)
        ws_dash[col].alignment = Alignment(horizontal="center")

    ws_dash["B5"] = "Đạt chuẩn"
    ws_dash["C5"] = f'=COUNTIF(Nhat_Ky_Chi_Tiet!F2:F{total_rows+1}, "ĐẠT")'
    ws_dash["D5"] = '=C5/C7'

    ws_dash["B6"] = "Mắc lỗi"
    ws_dash["C6"] = f'=COUNTIF(Nhat_Ky_Chi_Tiet!F2:F{total_rows+1}, "LỖI")'
    ws_dash["D6"] = '=C6/C7'

    ws_dash["B7"] = "Tổng số lượt"
    ws_dash["C7"] = '=SUM(C5:C6)'
    ws_dash["D7"] = '=SUM(D5:D6)'

    ws_dash["D5"].number_format = ws_dash["D6"].number_format = ws_dash["D7"].number_format = '0.0%'
    for r in range(4, 8):
        for c in ["B", "C", "D"]:
            ws_dash[f"{c}{r}"].border = thin_border

    pie_overall = PieChart()
    pie_overall.title = "Tỷ Lệ Đạt vs Mắc Lỗi"
    pie_overall.width, pie_overall.height = 11, 7
    pie_overall.add_data(Reference(ws_dash, min_col=3, min_row=4, max_row=6), titles_from_data=True)
    pie_overall.set_categories(Reference(ws_dash, min_col=2, min_row=5, max_row=6))
    pie_overall.dataLabels = DataLabelList()
    pie_overall.dataLabels.showPercent = True
    pie_overall.dataLabels.showVal = False
    ws_dash.add_chart(pie_overall, "B9")

    # Bảng phân tích chi tiết lỗi
    ws_dash["F4"], ws_dash["G4"], ws_dash["H4"] = "Phân loại lỗi thường gặp", "Số lần mắc", "Cơ cấu lỗi (%)"
    for col in ["F4", "G4", "H4"]:
        ws_dash[col].fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        ws_dash[col].font = Font(bold=True)
        ws_dash[col].alignment = Alignment(horizontal="center")

    error_list = sorted(list(error_types_set)) if error_types_set else ["Không phát hiện lỗi"]
    curr_row = 5
    for err in error_list:
        ws_dash[f"F{curr_row}"] = err
        if err == "Không phát hiện lỗi":
            ws_dash[f"G{curr_row}"] = ws_dash[f"H{curr_row}"] = 0
        else:
            ws_dash[f"G{curr_row}"] = f'=COUNTIF(Nhat_Ky_Chi_Tiet!G2:G{total_rows+1}, "{err}")'
            ws_dash[f"H{curr_row}"] = f'=IF($C$6>0, G{curr_row}/$C$6, 0)'
        ws_dash[f"H{curr_row}"].number_format = '0.0%'
        for c in ["F", "G", "H"]:
            ws_dash[f"{c}{curr_row}"].border = thin_border
        curr_row += 1

    total_err_row = curr_row
    ws_dash[f"F{total_err_row}"] = "Tổng số lỗi"
    ws_dash[f"G{total_err_row}"] = f'=SUM(G5:G{total_err_row-1})'
    ws_dash[f"H{total_err_row}"] = f'=SUM(H5:H{total_err_row-1})'
    ws_dash[f"H{total_err_row}"].number_format = '0.0%'
    for c in ["F", "G", "H"]:
        ws_dash[f"{c}{total_err_row}"].border = thin_border
        ws_dash[f"{c}{total_err_row}"].font = Font(bold=True)

    pie_error = PieChart()
    pie_error.title = "Cơ Cấu Các Dạng Lỗi Kỹ Thuật"
    pie_error.width, pie_error.height = 13, 7
    pie_error.add_data(Reference(ws_dash, min_col=7, min_row=4, max_row=total_err_row-1), titles_from_data=True)
    pie_error.set_categories(Reference(ws_dash, min_col=6, min_row=5, max_row=total_err_row-1))
    pie_error.dataLabels = DataLabelList()
    pie_error.dataLabels.showPercent = True
    pie_error.dataLabels.showVal = False
    ws_dash.add_chart(pie_error, "F9")

    for col in ["B", "C", "D", "F", "G", "H"]:
        ws_dash.column_dimensions[col].width = 24

    wb.save(output_filepath)
    return True, output_filepath

# ================= 3. GIAO DIỆN CHÍNH (TKINTER) =================
class SportsCoachApp:
    def __init__(self, window, window_title):
        self.window = window
        self.window.title(window_title)
        self.window.geometry("1180x720")
        self.window.configure(bg="#20232a")

        self.is_running = False
        self.cap = None
        self.current_stage = "CHUAN_BI"
        self.rep_count = 0
        self.last_capture_time = 0
        self.cooldown = 2.0
        self.session_records = []

        self.mp_pose = mp.solutions.pose
        self.pose = self.mp_pose.Pose(min_detection_confidence=0.6, min_tracking_confidence=0.6)
        self.mp_drawing = mp.solutions.drawing_utils

        self.filter_knee = AngleSmoother(5)
        self.filter_back = AngleSmoother(5)
        self.filter_elbow = AngleSmoother(5)
        self.filter_shoulder = AngleSmoother(5)

        self._build_ui()

    def _build_ui(self):
        left_panel = tk.Frame(self.window, width=340, bg="#282c34", padx=15, pady=15)
        left_panel.pack(side=tk.LEFT, fill=tk.Y)

        tk.Label(left_panel, text="TRỢ LÝ CHẤM KỸ THUẬT", font=("Helvetica", 14, "bold"), fg="#61afef", bg="#282c34").pack(pady=(0, 10))

        tk.Label(left_panel, text="1. Nội dung tập luyện:", font=("Helvetica", 10, "bold"), fg="#abb2bf", bg="#282c34").pack(anchor="w")
        self.exercise_var = tk.StringVar(value=list(EXERCISES.keys())[0])
        self.exercise_combo = ttk.Combobox(left_panel, textvariable=self.exercise_var, values=list(EXERCISES.keys()), state="readonly", width=32)
        self.exercise_combo.pack(pady=4, fill=tk.X)
        self.exercise_combo.bind("<<ComboboxSelected>>", self._on_exercise_change)

        self.lbl_desc = tk.Label(left_panel, text=EXERCISES[self.exercise_var.get()]["description"], 
                                 font=("Helvetica", 9, "italic"), fg="#98c379", bg="#282c34", wraplength=300, justify="left")
        self.lbl_desc.pack(pady=(0, 10), anchor="w")

        tk.Label(left_panel, text="2. Thông tin học sinh:", font=("Helvetica", 10, "bold"), fg="#abb2bf", bg="#282c34").pack(anchor="w")
        info_frame = tk.Frame(left_panel, bg="#282c34")
        info_frame.pack(fill=tk.X, pady=4)
        
        self.student_name = tk.StringVar(value="Nguyễn Văn A")
        self.txt_student = tk.Entry(info_frame, textvariable=self.student_name, font=("Helvetica", 10), bg="#3e4451", fg="white", width=20)
        self.txt_student.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        self.student_class = tk.StringVar(value="10A1")
        self.txt_class = tk.Entry(info_frame, textvariable=self.student_class, font=("Helvetica", 10), bg="#3e4451", fg="white", width=8)
        self.txt_class.pack(side=tk.RIGHT)

        stats_frame = tk.LabelFrame(left_panel, text=" Đánh giá trực tiếp ", font=("Helvetica", 10, "bold"), fg="#e5c07b", bg="#282c34", padx=10, pady=8)
        stats_frame.pack(fill=tk.X, pady=10)

        self.lbl_stat_phase = tk.Label(stats_frame, text="Pha: CHUẨN BỊ", font=("Helvetica", 10, "bold"), fg="#61afef", bg="#282c34")
        self.lbl_stat_phase.pack(anchor="w")

        self.lbl_stat_count = tk.Label(stats_frame, text="Số lần hoàn thành: 0", font=("Helvetica", 10), fg="#abb2bf", bg="#282c34")
        self.lbl_stat_count.pack(anchor="w", pady=2)

        self.lbl_stat_feedback = tk.Label(stats_frame, text="Chưa bật camera", font=("Helvetica", 10, "bold"), fg="#e06c75", bg="#282c34", wraplength=280)
        self.lbl_stat_feedback.pack(anchor="w", pady=2)

        self.btn_toggle_cam = tk.Button(left_panel, text="▶ BẬT CAMERA", font=("Helvetica", 11, "bold"), bg="#98c379", fg="#282c34", 
                                        relief=tk.FLAT, pady=6, command=self._toggle_camera)
        self.btn_toggle_cam.pack(fill=tk.X, pady=4)

        self.btn_export_excel = tk.Button(left_panel, text="📊 Xuất Báo Cáo Excel", font=("Helvetica", 10, "bold"), bg="#e5c07b", fg="#282c34",
                                          relief=tk.FLAT, pady=6, command=self._export_to_excel)
        self.btn_export_excel.pack(fill=tk.X, pady=4)

        self.btn_snap = tk.Button(left_panel, text="📷 Chụp ảnh tư thế này", font=("Helvetica", 9), bg="#4b5263", fg="white", 
                                  relief=tk.FLAT, pady=4, command=self._manual_capture)
        self.btn_snap.pack(fill=tk.X, pady=3)

        self.btn_open_folder = tk.Button(left_panel, text="📂 Mở thư mục ảnh lỗi", font=("Helvetica", 9), bg="#3e4451", fg="white", 
                                         relief=tk.FLAT, pady=4, command=self._open_folder)
        self.btn_open_folder.pack(fill=tk.X, pady=3)

        right_panel = tk.Frame(self.window, bg="#1e1e1e")
        right_panel.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH, padx=10, pady=10)

        self.video_label = tk.Label(right_panel, bg="#111", text="Vui lòng nhấn 'BẬT CAMERA' để bắt đầu", font=("Helvetica", 14), fg="#5c6370")
        self.video_label.pack(expand=True, fill=tk.BOTH)

    def _on_exercise_change(self, event=None):
        selected = self.exercise_var.get()
        self.lbl_desc.config(text=EXERCISES[selected]["description"])
        self.current_stage = "CHUAN_BI"
        self.rep_count = 0
        self.filter_knee.reset()
        self.filter_back.reset()
        self.filter_elbow.reset()
        self.filter_shoulder.reset()
        self._update_stats_ui("CHUẨN BỊ", 0, "Đã đổi nội dung bài tập", "#61afef")

    def _toggle_camera(self):
        if not self.is_running:
            self.cap = cv2.VideoCapture(0)
            if not self.cap.isOpened():
                messagebox.showerror("Lỗi Camera", "Không tìm thấy hoặc không thể mở Camera!")
                return
            self.is_running = True
            self.btn_toggle_cam.config(text="⏹ DỪNG CAMERA", bg="#e06c75", fg="white")
            self._update_loop()
        else:
            self.is_running = False
            if self.cap:
                self.cap.release()
            self.btn_toggle_cam.config(text="▶ BẬT CAMERA", bg="#98c379", fg="#282c34")
            self.video_label.config(image="", text="Camera đã dừng")

    def _update_stats_ui(self, phase, count, feedback, color):
        self.lbl_stat_phase.config(text=f"Pha: {phase}")
        self.lbl_stat_count.config(text=f"Số lần hoàn thành: {count}")
        self.lbl_stat_feedback.config(text=feedback, fg=color)

    def _record_performance(self, status, error_name, measured_angle, standard_note):
        record = {
            "Thời gian": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Họ tên": self.student_name.get().strip(),
            "Lớp": self.student_class.get().strip(),
            "Nội dung bài tập": self.exercise_var.get(),
            "Lần thứ": self.rep_count,
            "Kết quả": status,
            "Chi tiết lỗi": error_name,
            "Góc đo được": f"{measured_angle}°",
            "Tiêu chuẩn": standard_note
        }
        self.session_records.append(record)

    def _save_error_snapshot(self, frame, error_name, current_val, note):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        student = self.student_name.get().replace(" ", "_")
        filename = f"{OUTPUT_DIR}/{student}_{error_name}_{ts}.jpg"

        annotated = frame.copy()
        h, w, _ = annotated.shape
        cv2.rectangle(annotated, (10, h - 80), (w - 10, h - 10), (0, 0, 160), -1)
        cv2.putText(annotated, f"Hoc sinh: {self.student_name.get()} ({self.student_class.get()}) | Loi: {error_name}", (20, h - 50),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 2)
        cv2.putText(annotated, f"Goc: {current_val} deg | Chuan: {note} | {ts}", (20, h - 20),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0, 255, 255), 2)
        cv2.imwrite(filename, annotated)

    def _export_to_excel(self):
        if not self.session_records:
            messagebox.showwarning("Thông báo", "Chưa có dữ liệu tập luyện nào trong phiên này để xuất!")
            return

        default_filename = f"Bao_Cao_TDTT_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")],
            initialfile=default_filename
        )
        if not file_path:
            return

        try:
            success, msg = export_excel_with_chart(self.session_records, file_path)
            if success:
                res = messagebox.askyesno("Thành công", f"Đã xuất báo cáo Excel kèm biểu đồ tại:\n{file_path}\n\nBạn có muốn mở file ngay không?")
                if res:
                    os.system(f'start excel "{os.path.abspath(file_path)}"')
            else:
                messagebox.showerror("Lỗi", msg)
        except Exception as e:
            messagebox.showerror("Lỗi xuất file", f"Không thể lưu file: {str(e)}")

    def _manual_capture(self):
        if hasattr(self, 'current_frame') and self.current_frame is not None:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            student = self.student_name.get().replace(" ", "_")
            filename = f"{OUTPUT_DIR}/ChupThuCong_{student}_{ts}.jpg"
            cv2.imwrite(filename, self.current_frame)
            messagebox.showinfo("Thành công", f"Đã lưu ảnh tư thế hiện tại:\n{filename}")
        else:
            messagebox.showwarning("Cảnh báo", "Camera chưa được bật!")

    def _open_folder(self):
        os.system(f'explorer "{os.path.abspath(OUTPUT_DIR)}"' if os.name == 'nt' else f'open "{os.path.abspath(OUTPUT_DIR)}"')

    def _update_loop(self):
        if not self.is_running:
            return

        ret, frame = self.cap.read()
        if ret:
            self.current_frame = frame.copy()
            h, w, _ = frame.shape

            image_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            image_rgb.flags.writeable = False
            results = self.pose.process(image_rgb)
            image_rgb.flags.writeable = True

            current_time = time.time()
            ex_name = self.exercise_var.get()

            if results.pose_landmarks:
                lm = results.pose_landmarks.landmark

                # Mốc cơ thể bên trái (dùng cho Squat, Chạy)
                l_shoulder = [lm[self.mp_pose.PoseLandmark.LEFT_SHOULDER.value].x * w, lm[self.mp_pose.PoseLandmark.LEFT_SHOULDER.value].y * h]
                l_hip = [lm[self.mp_pose.PoseLandmark.LEFT_HIP.value].x * w, lm[self.mp_pose.PoseLandmark.LEFT_HIP.value].y * h]
                l_knee = [lm[self.mp_pose.PoseLandmark.LEFT_KNEE.value].x * w, lm[self.mp_pose.PoseLandmark.LEFT_KNEE.value].y * h]
                l_ankle = [lm[self.mp_pose.PoseLandmark.LEFT_ANKLE.value].x * w, lm[self.mp_pose.PoseLandmark.LEFT_ANKLE.value].y * h]

                # Mốc tay phải (dùng cho Cầu lông, Bóng rổ)
                r_shoulder = [lm[self.mp_pose.PoseLandmark.RIGHT_SHOULDER.value].x * w, lm[self.mp_pose.PoseLandmark.RIGHT_SHOULDER.value].y * h]
                r_elbow = [lm[self.mp_pose.PoseLandmark.RIGHT_ELBOW.value].x * w, lm[self.mp_pose.PoseLandmark.RIGHT_ELBOW.value].y * h]
                r_wrist = [lm[self.mp_pose.PoseLandmark.RIGHT_WRIST.value].x * w, lm[self.mp_pose.PoseLandmark.RIGHT_WRIST.value].y * h]
                r_hip = [lm[self.mp_pose.PoseLandmark.RIGHT_HIP.value].x * w, lm[self.mp_pose.PoseLandmark.RIGHT_HIP.value].y * h]

                if ex_name == "Squat / Bật nhảy bổ trợ":
                    knee_deg = self.filter_knee.update(calculate_angle(l_hip, l_knee, l_ankle))
                    back_deg = self.filter_back.update(calculate_angle(l_shoulder, l_hip, l_knee))

                    if self.current_stage == "CHUAN_BI":
                        if knee_deg > 160:
                            self._update_stats_ui("CHUẨN BỊ", self.rep_count, "Sẵn sàng: Hãy khuỵu gối tạo đà", "#61afef")
                        elif knee_deg <= 140:
                            self.current_stage = "TAO_DA"

                    elif self.current_stage == "TAO_DA":
                        if back_deg < EXERCISES[ex_name]["back_limit"]:
                            self._update_stats_ui("TẠO ĐÀ", self.rep_count, "LỖI: Lưng bị cúi gập quá nhiều!", "#e06c75")
                            if current_time - self.last_capture_time > self.cooldown:
                                self._save_error_snapshot(frame, "Cong_Lung", back_deg, f">={EXERCISES[ex_name]['back_limit']}")
                                self._record_performance("LỖI", "Cong_Lung", back_deg, f">={EXERCISES[ex_name]['back_limit']}°")
                                self.last_capture_time = current_time
                        elif knee_deg < 70:
                            self._update_stats_ui("TẠO ĐÀ", self.rep_count, "LỖI: Khuỵu gối quá sâu!", "#e06c75")
                            if current_time - self.last_capture_time > self.cooldown:
                                self._save_error_snapshot(frame, "Goi_Qua_Sau", knee_deg, "80 - 105")
                                self._record_performance("LỖI", "Goi_Qua_Sau", knee_deg, "80° - 105°")
                                self.last_capture_time = current_time
                        elif 80 <= knee_deg <= 105:
                            self._update_stats_ui("TẠO ĐÀ", self.rep_count, "Góc đẹp! Chuẩn bị bật/đứng lên", "#98c379")

                        if knee_deg > 150:
                            self.current_stage = "KET_THUC"

                    elif self.current_stage == "KET_THUC":
                        if knee_deg >= 160:
                            self.rep_count += 1
                            self._update_stats_ui("HOÀN THÀNH", self.rep_count, f"Tốt! Hoàn thành lần thứ {self.rep_count}", "#98c379")
                            self._record_performance("ĐẠT", "Chuẩn kỹ thuật", knee_deg, "Hoàn tất chu kỳ")
                            self.current_stage = "CHUAN_BI"

                    cv2.putText(frame, f"Goi: {knee_deg} deg", (int(l_knee[0]) + 15, int(l_knee[1])), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)

                elif ex_name == "Xuất phát thấp (Tư thế 'Sẵn sàng')":
                    knee_deg = self.filter_knee.update(calculate_angle(l_hip, l_knee, l_ankle))
                    cfg = EXERCISES[ex_name]
                    if cfg["front_knee_min"] <= knee_deg <= cfg["front_knee_max"]:
                        self._update_stats_ui("SẴN SÀNG", self.rep_count, "Góc gối chân trước đạt chuẩn (~90°)", "#98c379")
                    else:
                        self._update_stats_ui("SẴN SÀNG", self.rep_count, f"Chỉnh lại gối chân trước: {knee_deg}° (Chuẩn: 90°)", "#e06c75")
                    cv2.putText(frame, f"Goi truoc: {knee_deg} deg", (int(l_knee[0]) + 15, int(l_knee[1])), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)

                elif ex_name == "Cầu lông - Đập cầu thuận tay (Smash)":
                    arm_deg = self.filter_elbow.update(calculate_angle(r_shoulder, r_elbow, r_wrist))
                    armpit_deg = self.filter_shoulder.update(calculate_angle(r_hip, r_shoulder, r_elbow))

                    if armpit_deg < 85:
                        self._update_stats_ui("GIƯƠNG CUNG", self.rep_count, "LỖI: Khuỷu tay hạ quá thấp (khép nách)!", "#e06c75")
                        if current_time - self.last_capture_time > self.cooldown:
                            self._save_error_snapshot(frame, "Khep_Nach_Dap_Cau", armpit_deg, ">= 90°")
                            self._record_performance("LỖI", "Khep_Nach_Dap_Cau", armpit_deg, ">= 90°")
                            self.last_capture_time = current_time
                    elif arm_deg >= 165:
                        self._update_stats_ui("TIẾP XÚC CẦU", self.rep_count, "Tốt! Duỗi tay ở tầm cao nhất", "#98c379")
                    else:
                        self._update_stats_ui("THEO DÕI", self.rep_count, "Quan sát mở vợt...", "#61afef")

                    cv2.putText(frame, f"Khuyu tay: {arm_deg} deg", (int(r_elbow[0]) + 15, int(r_elbow[1])), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)

                elif ex_name == "Bóng rổ - Ném rổ 1 tay trên vai":
                    arm_deg = self.filter_elbow.update(calculate_angle(r_shoulder, r_elbow, r_wrist))
                    if 85 <= arm_deg <= 98:
                        self._update_stats_ui("HẠ ĐÀ / KHÓA GÓC", self.rep_count, "Tốt! Khuỷu tay khóa chuẩn góc vuông 90°", "#98c379")
                    elif arm_deg >= 165:
                        self._update_stats_ui("RA BÓNG", self.rep_count, "Duỗi thẳng tay ném hoàn toàn!", "#98c379")
                    else:
                        self._update_stats_ui("CHỜ NÉM", self.rep_count, "Giữ góc khuỷu 90°", "#61afef")

                    cv2.putText(frame, f"Khuyu tay: {arm_deg} deg", (int(r_elbow[0]) + 15, int(r_elbow[1])), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)

                self.mp_drawing.draw_landmarks(frame, results.pose_landmarks, self.mp_pose.POSE_CONNECTIONS)
            else:
                self.filter_knee.reset()
                self.filter_back.reset()
                self.filter_elbow.reset()
                self.filter_shoulder.reset()

            cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)
            img = Image.fromarray(cv2image)
            imgtk = ImageTk.PhotoImage(image=img)
            self.video_label.imgtk = imgtk
            self.video_label.configure(image=imgtk)

        self.window.after(15, self._update_loop)

if __name__ == "__main__":
    root = tk.Tk()
    app = SportsCoachApp(root, "Phần Mềm Hỗ Trợ Giảng Dạy & Chỉnh Sửa Kỹ Thuật Thể Dục Thể Thao")
    root.mainloop()