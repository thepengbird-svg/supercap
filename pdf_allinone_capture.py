#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
올인원 PDF 캡처 & 크롭 프로그램
사용자 친화적 GUI로 캡처부터 PDF 생성까지 완전 자동화
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pyautogui
import cv2
import numpy as np
import os
import time
from PIL import Image
from datetime import datetime
import threading

class PDFCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("올인원 PDF 캡처 & 크롭 프로그램 v1.0")
        self.root.geometry("600x700")
        self.root.resizable(False, False)
        
        # 변수 초기화
        self.total_pages = tk.IntVar(value=200)
        self.page_key = tk.StringVar(value="pagedown")
        self.output_folder = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self.crop_area = None
        self.is_capturing = False
        self.is_paused = False
        self.current_page = 0
        
        self.setup_ui()
        
    def setup_ui(self):
        """UI 구성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="PDF 캡처 & 크롭 프로그램", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 설정 섹션
        settings_frame = ttk.LabelFrame(main_frame, text="설정", padding="10")
        settings_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # 페이지 수 설정
        ttk.Label(settings_frame, text="총 페이지 수:").grid(row=0, column=0, sticky=tk.W, pady=5)
        page_spinbox = ttk.Spinbox(settings_frame, from_=1, to=2000, width=10, 
                                  textvariable=self.total_pages)
        page_spinbox.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # 페이지 넘기기 키 설정
        ttk.Label(settings_frame, text="페이지 넘기기 키:").grid(row=1, column=0, sticky=tk.W, pady=5)
        key_frame = ttk.Frame(settings_frame)
        key_frame.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        
        # 키 선택 라디오 버튼
        keys = [("Page Down", "pagedown"), ("스페이스바", "space"), 
                ("엔터", "enter"), ("오른쪽 화살표", "right"), ("아래 화살표", "down")]
        
        for i, (text, value) in enumerate(keys):
            ttk.Radiobutton(key_frame, text=text, variable=self.page_key, 
                           value=value).grid(row=i//3, column=i%3, sticky=tk.W, padx=5)
        
        # 저장 위치 설정
        ttk.Label(settings_frame, text="저장 위치:").grid(row=3, column=0, sticky=tk.W, pady=5)
        folder_frame = ttk.Frame(settings_frame)
        folder_frame.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        
        ttk.Entry(folder_frame, textvariable=self.output_folder, width=30).grid(row=0, column=0)
        ttk.Button(folder_frame, text="찾아보기", command=self.choose_folder).grid(row=0, column=1, padx=(5, 0))
        
        # 진행 단계 표시
        progress_frame = ttk.LabelFrame(main_frame, text="진행 단계", padding="10")
        progress_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 20))
        
        self.step_labels = []
        steps = ["1단계: 설정 완료", "2단계: 크롭 영역 선택", "3단계: 캡처 진행", 
                "4단계: 크롭 & PDF 생성", "5단계: 완료"]
        
        for i, step in enumerate(steps):
            label = ttk.Label(progress_frame, text=step, foreground="gray")
            label.grid(row=i, column=0, sticky=tk.W, pady=2)
            self.step_labels.append(label)
        
        # 진행률 바
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                          maximum=100, length=400)
        self.progress_bar.grid(row=len(steps), column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        self.progress_label = ttk.Label(progress_frame, text="준비 완료")
        self.progress_label.grid(row=len(steps)+1, column=0, pady=5)
        
        # 버튼 섹션
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        self.start_button = ttk.Button(button_frame, text="시작하기", 
                                      command=self.start_process, style="Accent.TButton")
        self.start_button.grid(row=0, column=0, padx=5)
        
        self.pause_button = ttk.Button(button_frame, text="일시정지", 
                                      command=self.pause_process, state="disabled")
        self.pause_button.grid(row=0, column=1, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="중단", 
                                     command=self.stop_process, state="disabled")
        self.stop_button.grid(row=0, column=2, padx=5)
        
        # 로그 섹션
        log_frame = ttk.LabelFrame(main_frame, text="진행 로그", padding="10")
        log_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # 스크롤 가능한 텍스트 위젯
        self.log_text = tk.Text(log_frame, height=8, width=60, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 그리드 가중치 설정
        main_frame.columnconfigure(1, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log("프로그램이 시작되었습니다. 설정을 확인하고 '시작하기'를 클릭하세요.")
    
    def log(self, message):
        """로그 메시지 추가"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.root.update()
    
    def update_step(self, step_index, status="active"):
        """진행 단계 업데이트"""
        colors = {"active": "blue", "completed": "green", "pending": "gray"}
        
        for i, label in enumerate(self.step_labels):
            if i < step_index:
                label.configure(foreground=colors["completed"])
            elif i == step_index:
                label.configure(foreground=colors[status])
            else:
                label.configure(foreground=colors["pending"])
    
    def update_progress(self, value, text):
        """진행률 업데이트"""
        self.progress_var.set(value)
        self.progress_label.configure(text=text)
        self.root.update()
    
    def choose_folder(self):
        """저장 폴더 선택"""
        folder = filedialog.askdirectory(initialdir=self.output_folder.get())
        if folder:
            self.output_folder.set(folder)

    def start_process(self):
        """전체 프로세스 시작"""
        self.log("프로세스를 시작합니다...")
        self.update_step(0, "completed")
        
        # 버튼 상태 변경
        self.start_button.configure(state="disabled")
        self.pause_button.configure(state="normal")
        self.stop_button.configure(state="normal")
        
        # 단계별 진행
        if self.prepare_capture():
            if self.select_crop_area():
                self.start_capture()
    
    def prepare_capture(self):
        """캡처 준비"""
        self.log("캡처 준비 중...")
        
        # 사용자 안내
        result = messagebox.askyesno(
            "캡처 준비",
            "캡처를 시작하기 전에 다음을 확인해주세요:\n\n"
            "1. 캡처할 문서를 열어주세요\n"
            "2. 첫 번째 페이지로 이동해주세요\n"
            "3. 문서 창을 최대한 크게 키워주세요\n"
            "4. 다른 창들을 최소화해주세요\n\n"
            "준비가 완료되었나요?"
        )
        
        if not result:
            self.reset_ui()
            return False
        
        self.log("캡처 준비 완료!")
        return True
    
    def select_crop_area(self):
        """크롭 영역 선택"""
        self.log("크롭 영역 선택을 시작합니다...")
        self.update_step(1, "active")
        
        # 첫 번째 스크린샷 촬영
        time.sleep(2)  # 사용자가 준비할 시간
        screenshot = pyautogui.screenshot()
        temp_image_path = os.path.join(self.output_folder.get(), "temp_screenshot.png")
        screenshot.save(temp_image_path)
        
        # 크롭 영역 선택
        self.crop_area = self.select_crop_area_manually(temp_image_path)
        
        if self.crop_area is None:
            self.log("크롭 영역 선택이 취소되었습니다.")
            self.reset_ui()
            return False
        
        # 임시 파일 삭제
        try:
            os.remove(temp_image_path)
        except:
            pass
        
        self.log(f"크롭 영역 선택 완료: {self.crop_area}")
        self.update_step(1, "completed")
        return True
    
    def select_crop_area_manually(self, image_path):
        """사용자가 마우스로 크롭 영역을 직접 선택"""
        self.log("마우스로 PDF 페이지 영역을 선택해주세요...")
        
        # 이미지 읽기
        img = cv2.imread(image_path)
        if img is None:
            self.log("스크린샷을 읽을 수 없습니다!")
            return None
        
        # 화면 크기에 맞게 조정
        height, width = img.shape[:2]
        max_display_width = 1200
        max_display_height = 800
        
        if width > max_display_width or height > max_display_height:
            scale_w = max_display_width / width
            scale_h = max_display_height / height
            scale = min(scale_w, scale_h)
            
            new_width = int(width * scale)
            new_height = int(height * scale)
            
            display_img = cv2.resize(img, (new_width, new_height))
            scale_factor = scale
        else:
            display_img = img.copy()
            scale_factor = 1.0
        
        # 안내 메시지
        messagebox.showinfo(
            "크롭 영역 선택",
            "마우스로 PDF 페이지 영역을 드래그하여 선택하세요!\n\n"
            "• 마우스 드래그로 영역 선택\n"
            "• SPACE 또는 ENTER: 선택 완료\n"
            "• R: 다시 선택\n"
            "• ESC: 취소"
        )
        
        # ROI 선택
        roi = cv2.selectROI("크롭 영역 선택 - PDF 페이지를 드래그하세요", 
                           display_img, showCrosshair=True)
        cv2.destroyAllWindows()
        
        x, y, w, h = roi
        if w <= 0 or h <= 0:
            return None
        
        # 원본 크기로 변환
        if scale_factor != 1.0:
            x = int(x / scale_factor)
            y = int(y / scale_factor)
            w = int(w / scale_factor)
            h = int(h / scale_factor)
        
        # 범위 조정
        orig_height, orig_width = img.shape[:2]
        x = max(0, min(x, orig_width - 1))
        y = max(0, min(y, orig_height - 1))
        w = min(w, orig_width - x)
        h = min(h, orig_height - y)
        
        # 확인
        confirm = messagebox.askyesno(
            "크롭 영역 확인",
            f"선택된 크롭 영역:\n"
            f"X: {x}, Y: {y}\n"
            f"너비: {w}, 높이: {h}\n\n"
            f"이 영역으로 진행하시겠습니까?"
        )
        
        return (x, y, w, h) if confirm else None
    
    def start_capture(self):
        """캡처 시작"""
        self.log("자동 캡처를 시작합니다...")
        self.update_step(2, "active")
        self.is_capturing = True
        
        # 백그라운드 쓰레드에서 캡처 실행
        capture_thread = threading.Thread(target=self.capture_pages)
        capture_thread.daemon = True
        capture_thread.start()
    
    def capture_pages(self):
        """페이지 캡처 실행"""
        try:
            # 저장 폴더 생성
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.capture_folder = os.path.join(self.output_folder.get(), f"captured_pages_{timestamp}")
            os.makedirs(self.capture_folder, exist_ok=True)
            
            self.log(f"캡처 폴더 생성: {self.capture_folder}")
            
            # 캡처 시작 전 잠시 대기
            time.sleep(3)
            
            success_count = 0
            total_pages = self.total_pages.get()
            
            for page in range(1, total_pages + 1):
                if not self.is_capturing:
                    self.log("캡처가 중단되었습니다.")
                    break
                
                # 일시정지 체크
                while self.is_paused and self.is_capturing:
                    time.sleep(0.1)
                
                if not self.is_capturing:
                    break
                
                self.current_page = page
                
                # 현재 페이지 캡처
                self.log(f"페이지 {page}/{total_pages} 캡처 중...")
                
                screenshot = pyautogui.screenshot()
                filename = f"page_{page:03d}.png"
                filepath = os.path.join(self.capture_folder, filename)
                screenshot.save(filepath)
                
                success_count += 1
                
                # 진행률 업데이트
                progress = (page / total_pages) * 50  # 50%까지가 캡처
                self.update_progress(progress, f"캡처 진행: {page}/{total_pages}")
                
                # 마지막 페이지가 아니면 다음 페이지로
                if page < total_pages and self.is_capturing:
                    self.next_page()
                    time.sleep(1.5)  # 페이지 로딩 대기
            
            if self.is_capturing:
                self.log(f"캡처 완료! {success_count}/{total_pages} 페이지 성공")
                self.update_step(2, "completed")
                self.start_crop_and_pdf()
            
        except Exception as e:
            self.log(f"캡처 중 오류 발생: {e}")
            self.reset_ui()
    
    def next_page(self):
        """다음 페이지로 이동"""
        key = self.page_key.get()
        pyautogui.press(key)
    
    def start_crop_and_pdf(self):
        """크롭 및 PDF 생성"""
        self.log("크롭 및 PDF 생성을 시작합니다...")
        self.update_step(3, "active")
        
        # 백그라운드에서 실행
        crop_thread = threading.Thread(target=self.crop_and_create_pdf)
        crop_thread.daemon = True
        crop_thread.start()
    
    def crop_and_create_pdf(self):
        """크롭 및 PDF 생성 실행"""
        try:
            # 크롭 폴더 생성
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.cropped_folder = os.path.join(self.output_folder.get(), f"cropped_pages_{timestamp}")
            os.makedirs(self.cropped_folder, exist_ok=True)
            
            # 캡처된 이미지들 가져오기
            image_files = [f for f in os.listdir(self.capture_folder) if f.endswith('.png')]
            image_files.sort()
            
            if not image_files:
                self.log("캡처된 이미지가 없습니다!")
                return
            
            self.log(f"총 {len(image_files)}개 이미지 크롭 중...")
            
            # 모든 이미지 크롭
            success_count = 0
            for i, img_file in enumerate(image_files, 1):
                if not self.is_capturing:
                    break
                
                input_path = os.path.join(self.capture_folder, img_file)
                output_path = os.path.join(self.cropped_folder, img_file)
                
                if self.crop_image(input_path, self.crop_area, output_path):
                    success_count += 1
                
                # 진행률 업데이트 (50-90%)
                progress = 50 + (i / len(image_files)) * 40
                self.update_progress(progress, f"크롭 진행: {i}/{len(image_files)}")
            
            if not self.is_capturing:
                return
            
            self.log(f"크롭 완료! {success_count}/{len(image_files)} 성공")
            
            # PDF 생성
            self.log("PDF 생성 중...")
            output_pdf = os.path.join(self.output_folder.get(), f"final_book_{timestamp}.pdf")
            
            if self.create_pdf_from_images(self.cropped_folder, output_pdf):
                self.update_progress(100, "완료!")
                self.update_step(4, "completed")
                self.log(f"모든 작업 완료! PDF 생성: {output_pdf}")
                
                # 완료 처리
                self.root.after(0, lambda: self.process_completed(output_pdf))
            else:
                self.log("PDF 생성 실패!")
                
        except Exception as e:
            self.log(f"크롭/PDF 생성 중 오류: {e}")
            self.reset_ui()
    
    def crop_image(self, image_path, crop_area, output_path):
        """이미지 크롭"""
        try:
            img = Image.open(image_path)
            x, y, w, h = crop_area
            cropped = img.crop((x, y, x + w, y + h))
            cropped.save(output_path)
            return True
        except Exception as e:
            self.log(f"크롭 실패 {image_path}: {e}")
            return False
    
    def create_pdf_from_images(self, folder_path, output_pdf):
        """이미지들을 PDF로 변환"""
        try:
            image_files = [f for f in os.listdir(folder_path) if f.endswith('.png')]
            image_files.sort()
            
            if not image_files:
                return False
            
            # 첫 번째 이미지로 PDF 생성
            first_image_path = os.path.join(folder_path, image_files[0])
            first_image = Image.open(first_image_path)
            
            if first_image.mode != 'RGB':
                first_image = first_image.convert('RGB')
            
            # 나머지 이미지들
            other_images = []
            for img_file in image_files[1:]:
                img_path = os.path.join(folder_path, img_file)
                img = Image.open(img_path)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                other_images.append(img)
            
            # PDF 저장
            first_image.save(output_pdf, save_all=True, append_images=other_images, 
                           resolution=200.0, quality=95)
            
            return True
            
        except Exception as e:
            self.log(f"PDF 생성 실패: {e}")
            return False

    def pause_process(self):
        """프로세스 일시정지/재개"""
        if self.is_paused:
            self.is_paused = False
            self.pause_button.configure(text="일시정지")
            self.log("캡처 재개됨")
        else:
            self.is_paused = True
            self.pause_button.configure(text="재개")
            self.log("캡처 일시정지됨")
    
    def stop_process(self):
        """프로세스 중단"""
        result = messagebox.askyesno("중단 확인", "정말로 작업을 중단하시겠습니까?")
        if result:
            self.is_capturing = False
            self.is_paused = False
            self.log("작업이 중단되었습니다.")
            self.reset_ui()
    
    def reset_ui(self):
        """UI 초기화"""
        self.start_button.configure(state="normal")
        self.pause_button.configure(state="disabled", text="일시정지")
        self.stop_button.configure(state="disabled")
        self.progress_var.set(0)
        self.progress_label.configure(text="준비 완료")
        self.update_step(-1, "pending")
        self.is_capturing = False
        self.is_paused = False
    
    def process_completed(self, output_pdf):
        """프로세스 완료 처리"""
        self.reset_ui()
        self.update_step(4, "completed")
        
        # 완료 메시지
        result = messagebox.askyesno(
            "작업 완료!",
            f"PDF 생성이 완료되었습니다!\n\n"
            f"파일 위치: {output_pdf}\n\n"
            f"PDF 파일을 바로 열어보시겠습니까?"
        )
        
        if result:
            try:
                os.startfile(output_pdf)
            except Exception as e:
                self.log(f"PDF 파일 열기 실패: {e}")
        
        self.log("모든 작업이 완료되었습니다!")

def main():
    """메인 실행 함수"""
    # 필요한 라이브러리 확인
    try:
        import pyautogui
        import cv2
        import PIL
    except ImportError as e:
        missing_lib = str(e).split("'")[1] if "'" in str(e) else "알 수 없는 라이브러리"
        
        root = tk.Tk()
        root.withdraw()
        
        messagebox.showerror(
            "라이브러리 누락",
            f"필요한 라이브러리가 설치되지 않았습니다: {missing_lib}\n\n"
            f"다음 명령어로 설치해주세요:\n"
            f"pip install pyautogui opencv-python pillow"
        )
        return
    
    # GUI 실행
    root = tk.Tk()
    app = PDFCaptureApp(root)
    
    # 종료 처리
    def on_closing():
        if app.is_capturing:
            result = messagebox.askyesno(
                "종료 확인", 
                "작업이 진행 중입니다. 정말로 종료하시겠습니까?"
            )
            if result:
                app.is_capturing = False
                root.destroy()
        else:
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 환영 메시지
    messagebox.showinfo(
        "환영합니다!",
        "올인원 PDF 캡처 & 크롭 프로그램에 오신 것을 환영합니다!\n\n"
        "이 프로그램으로 다음을 한 번에 할 수 있습니다:\n"
        "• 문서 자동 캡처\n"
        "• 원하는 영역 크롭\n"
        "• 고품질 PDF 생성\n\n"
        "사용법:\n"
        "1. 설정을 확인하세요\n"
        "2. '시작하기' 버튼을 클릭하세요\n"
        "3. 안내에 따라 진행하세요"
    )
    
    root.mainloop()

if __name__ == "__main__":
    main()
