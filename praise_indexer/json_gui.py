#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JSON 기반 찬양 검색 GUI (CustomTkinter)
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import json
from pathlib import Path
import threading
import time
import sys
import os

from json_indexer import JSONPraiseIndexer
from json_ppt_generator_fixed import JSONPPTGeneratorFixed as JSONPPTGenerator

class JSONPraiseGUI:
    """JSON 기반 찬양 검색 GUI"""
    
    def __init__(self):
        # 리소스 경로 헬퍼: 실행파일과 같은 폴더의 파일을 찾음
        def resource_path(relative: str) -> Path:
            if getattr(sys, 'frozen', False):
                # PyInstaller로 빌드된 실행파일인 경우
                base_path = Path(sys.executable).parent
            else:
                # 개발 환경인 경우
                base_path = Path(__file__).parent
            return (base_path / relative).resolve()
        self.resource_path = resource_path

        # CustomTkinter 설정
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.root = ctk.CTk()
        self.root.title("찬양 검색 프로그램 (JSON 버전)")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)
        
        # 데이터
        # JSON/템플릿 경로를 실행파일과 같은 폴더 기준으로 설정
        self.json_path = str(self.resource_path("praise_index.json"))
        self.template_path = str(self.resource_path("temp.pptx"))

        # 중복 라인 보존: remove_duplicate_lines=False
        self.indexer = JSONPraiseIndexer(output_json=self.json_path, remove_duplicate_lines=False)
        self.generator = None
        self.search_results = []
        self.selected_praises = []
        self.selected_indices = set()  # 선택된 항목들의 인덱스
        self.dragging_index = None  # 드래그 시작 인덱스 (순서 변경용)
        self.drop_indicator = None  # 드롭 위치 표시선
        self.drag_start_y = None  # 드래그 시작 Y 좌표
        
        # 검색 타이머
        self.search_timer = None
        
        self.setup_ui()
        self.load_data()
    
    def create_tooltip(self, widget, text):
        """툴팁 생성"""
        def on_enter(event):
            tooltip = ctk.CTkToplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = ctk.CTkLabel(tooltip, text=text, 
                               font=ctk.CTkFont(size=12),
                               fg_color=("gray90", "gray20"),
                               corner_radius=5)
            label.pack(padx=5, pady=5)
            
            widget.tooltip = tooltip
        
        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
    
    def setup_ui(self):
        """UI 설정"""
        # 메인 프레임
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 제목
        title_label = ctk.CTkLabel(main_frame, text="찬양 검색 프로그램", 
                                  font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(pady=(0, 10))
        
        # 검색 프레임
        search_frame = ctk.CTkFrame(main_frame)
        search_frame.pack(fill="x", pady=(0, 10))
        
        # 검색 입력
        search_label = ctk.CTkLabel(search_frame, text="검색어:", font=ctk.CTkFont(size=14, weight="bold"))
        search_label.pack(side="left", padx=(10, 5), pady=10)
        
        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_var, 
                                       width=300, height=30, font=ctk.CTkFont(size=14))
        self.search_entry.pack(side="left", padx=(0, 10), pady=10)
        self.search_entry.bind('<KeyRelease>', self.on_search_change)
        self.search_entry.bind('<Return>', self.on_search_enter)
        
        # 검색 타입
        search_type_label = ctk.CTkLabel(search_frame, text="검색 타입:", font=ctk.CTkFont(size=13))
        search_type_label.pack(side="left", padx=(10, 5), pady=10)
        
        self.search_type_var = tk.StringVar(value="both")
        search_type_menu = ctk.CTkOptionMenu(search_frame, variable=self.search_type_var,
                                           values=["제목", "가사", "전체"], width=80, height=30,
                                           font=ctk.CTkFont(size=13))
        search_type_menu.pack(side="left", padx=(0, 10), pady=10)
        
        # 파일 추가 버튼
        self.add_file_button = ctk.CTkButton(search_frame, text="파일 추가", command=self.add_pptx_file,
                                           width=80, height=30, font=ctk.CTkFont(size=13, weight="bold"),
                                           fg_color="green", hover_color="darkgreen")
        self.add_file_button.pack(side="right", padx=(0, 10), pady=10)
        
        # 인덱싱 버튼
        self.index_button = ctk.CTkButton(search_frame, text="인덱싱", command=self.reindex_data,
                                         width=70, height=30, font=ctk.CTkFont(size=13, weight="bold"))
        self.index_button.pack(side="right", padx=(0, 10), pady=10)
        
        # 메인 콘텐츠 프레임
        content_frame = ctk.CTkFrame(main_frame)
        content_frame.pack(fill="both", expand=True)
        
        # 왼쪽 프레임 (검색 결과)
        left_frame = ctk.CTkFrame(content_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # 검색 결과 제목
        results_title = ctk.CTkLabel(left_frame, text="검색 결과", 
                                   font=ctk.CTkFont(size=16, weight="bold"))
        results_title.pack(pady=(10, 5))
        
        # 검색 결과 리스트
        self.results_frame = ctk.CTkScrollableFrame(left_frame, height=300)
        self.results_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # 오른쪽 프레임 (선택된 찬양)
        right_frame = ctk.CTkFrame(content_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # 선택된 찬양 제목
        selected_title = ctk.CTkLabel(right_frame, text="선택된 찬양", 
                                    font=ctk.CTkFont(size=16, weight="bold"))
        selected_title.pack(pady=(10, 5))
        
        # 선택된 찬양 리스트
        self.selected_frame = ctk.CTkScrollableFrame(right_frame, height=300)
        self.selected_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # 선택된 찬양 버튼들
        button_frame = ctk.CTkFrame(right_frame)
        button_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        # 파일 추가 버튼
        self.add_file_button = ctk.CTkButton(button_frame, text="파일 추가", command=self.add_pptx_file, 
                                           width=80, height=30, fg_color="green", hover_color="darkgreen",
                                           font=ctk.CTkFont(size=12))
        self.add_file_button.pack(side="left", padx=(0, 3))
        
        # 삭제 버튼 (선택된 항목들)
        self.delete_selected_button = ctk.CTkButton(button_frame, text="선택 삭제", command=self.delete_selected_items, 
                                                   width=80, height=30, fg_color="red", hover_color="darkred",
                                                   font=ctk.CTkFont(size=12))
        self.delete_selected_button.pack(side="left", padx=(0, 3))
        
        # 전체 선택/해제 버튼
        self.select_all_button = ctk.CTkButton(button_frame, text="전체 선택", command=self.select_all_items, 
                                             width=80, height=30, fg_color="blue", hover_color="darkblue",
                                             font=ctk.CTkFont(size=12))
        self.select_all_button.pack(side="left", padx=(0, 3))
        
        # 전체 제거 버튼
        ctk.CTkButton(button_frame, text="전체 제거", command=self.clear_selected, 
                     width=80, height=30, font=ctk.CTkFont(size=12)).pack(side="left", padx=(0, 3))
        
        # PPT 생성 프레임
        ppt_frame = ctk.CTkFrame(main_frame)
        ppt_frame.pack(fill="x", pady=(10, 0))
        
        # PPT 생성 버튼
        self.ppt_button = ctk.CTkButton(ppt_frame, text="PPT 생성", command=self.generate_ppt,
                                      width=100, height=35, font=ctk.CTkFont(size=16, weight="bold"))
        self.ppt_button.pack(side="left", padx=10, pady=10)
        
        # 진행률 표시
        self.progress_var = tk.StringVar(value="준비됨")
        self.progress_label = ctk.CTkLabel(ppt_frame, textvariable=self.progress_var, 
                                         font=ctk.CTkFont(size=14))
        self.progress_label.pack(side="left", padx=(10, 0), pady=10)
    
    def load_data(self):
        """데이터 로드"""
        try:
            if self.indexer.load_from_json():
                self.generator = JSONPPTGenerator(
                    json_file=self.json_path,
                    template_file=self.template_path
                )
                self.progress_var.set(f"로드됨: {len(self.indexer.praise_data)}개 찬양")
            else:
                self.progress_var.set("JSON 파일이 없습니다. 인덱싱을 실행하세요.")
        except Exception as e:
            messagebox.showerror("오류", f"데이터 로드 실패: {e}")
    
    def reindex_data(self):
        """데이터 재인덱싱"""
        def index_thread():
            try:
                self.progress_var.set("인덱싱 중...")
                self.root.update()
                
                success = self.indexer.index_praise_files()
                
                if success:
                    self.generator = JSONPPTGenerator(
                        json_file=self.json_path,
                        template_file=self.template_path
                    )
                    self.progress_var.set(f"인덱싱 완료: {len(self.indexer.praise_data)}개 찬양")
                    messagebox.showinfo("완료", "인덱싱이 완료되었습니다.")
                else:
                    self.progress_var.set("인덱싱 실패")
                    messagebox.showerror("오류", "인덱싱에 실패했습니다.")
            except Exception as e:
                self.progress_var.set("인덱싱 실패")
                messagebox.showerror("오류", f"인덱싱 실패: {e}")
        
        threading.Thread(target=index_thread, daemon=True).start()
    
    def on_search_change(self, event):
        """검색어 변경 시"""
        if self.search_timer:
            self.root.after_cancel(self.search_timer)
        
        # 디바운싱 시간을 100ms로 단축 (매우 빠른 반응성)
        self.search_timer = self.root.after(100, self.perform_search)
    
    def on_search_enter(self, event):
        """엔터키로 검색"""
        if self.search_timer:
            self.root.after_cancel(self.search_timer)
        
        self.perform_search()
    
    def perform_search(self):
        """검색 수행"""
        query = self.search_var.get().strip()
        search_type = self.search_type_var.get()
        
        if not query:
            self.search_results = []
            self.update_results_display()
            return
        
        # 최소 검색어 길이 체크 (성능 개선)
        if len(query) < 2:
            self.search_results = []
            self.update_results_display()
            return
        
        try:
            # 검색 타입 변환
            type_map = {"제목": "title", "가사": "lyrics", "전체": "both"}
            search_type = type_map.get(search_type, "both")
            
            # 검색 결과 제한 (성능 개선)
            self.search_results = self.indexer.search_praises(query, search_type)[:30]  # 최대 30개로 제한
            self.update_results_display()
        except Exception as e:
            messagebox.showerror("오류", f"검색 실패: {e}")
    
    def update_results_display(self):
        """검색 결과 표시 업데이트"""
        # 기존 위젯 제거
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        if not self.search_results:
            no_results_label = ctk.CTkLabel(self.results_frame, text="검색 결과가 없습니다.", 
                                           font=ctk.CTkFont(size=14))
            no_results_label.pack(pady=20)
            return
        
        # 새 결과 추가
        for i, praise in enumerate(self.search_results):
            self.create_result_item(praise, i)
    
    def create_result_item(self, praise, index):
        """검색 결과 아이템 생성"""
        # 메인 프레임
        item_frame = ctk.CTkFrame(self.results_frame)
        item_frame.pack(fill="x", padx=3, pady=3)
        
        # 제목
        title_label = ctk.CTkLabel(item_frame, text=praise['title'], 
                                 font=ctk.CTkFont(size=15, weight="bold"),
                                 anchor="w")
        title_label.pack(fill="x", padx=10, pady=(8, 3))
        
        # 슬라이드 수
        slides_count = len(praise.get('slides_text', []))
        slides_label = ctk.CTkLabel(item_frame, text=f"슬라이드: {slides_count}개", 
                                  font=ctk.CTkFont(size=12), text_color="gray")
        slides_label.pack(anchor="w", padx=10, pady=(0, 3))
        
        # 가사 미리보기
        lyrics_preview = self.get_lyrics_preview(praise)
        if lyrics_preview:
            lyrics_label = ctk.CTkLabel(item_frame, text=lyrics_preview, 
                                      font=ctk.CTkFont(size=12), 
                                      text_color="lightblue", anchor="w")
            lyrics_label.pack(fill="x", padx=10, pady=(0, 8))
        
        # 선택 버튼
        select_button = ctk.CTkButton(item_frame, text="선택", command=lambda: self.add_to_selected(praise),
                                    width=60, height=25, font=ctk.CTkFont(size=12))
        select_button.pack(anchor="e", padx=10, pady=(0, 8))
    
    def get_lyrics_preview(self, praise):
        """가사 미리보기 생성"""
        try:
            slides_text = praise.get('slides_text', [])
            if slides_text and len(slides_text) > 0:
                # 첫 번째 슬라이드의 첫 번째 줄들
                first_slide = slides_text[0]
                lines = first_slide.get('text_lines', [])
                if lines:
                    preview = lines[0] if lines else ""
                    if len(lines) > 1:
                        preview += f" ... {lines[1]}" if lines[1] else ""
                    return preview[:100] + "..." if len(preview) > 100 else preview
            return ""
        except:
            return ""
    
    def add_to_selected(self, praise):
        """선택된 찬양에 추가"""
        # 중복 확인
        for selected in self.selected_praises:
            if selected['id'] == praise['id']:
                return
        
        self.selected_praises.append(praise)
        self.update_selected_display()
    
    def update_selected_display(self):
        """선택된 찬양 표시 업데이트"""
        # 기존 위젯 제거
        for widget in self.selected_frame.winfo_children():
            widget.destroy()
        
        if not self.selected_praises:
            no_selected_label = ctk.CTkLabel(self.selected_frame, text="선택된 찬양이 없습니다.", 
                                           font=ctk.CTkFont(size=14))
            no_selected_label.pack(pady=20)
            return
        
        # 새 선택된 찬양 추가
        for i, praise in enumerate(self.selected_praises):
            self.create_selected_item(praise, i)
        
        # 체크박스 상태 업데이트
        self.update_checkbox_states()
    
    def create_selected_item(self, praise, index):
        """선택된 찬양 아이템 생성 (체크박스 선택 및 드래그-앤-드롭 정렬)"""
        # 메인 프레임
        item_frame = ctk.CTkFrame(self.selected_frame)
        item_frame.pack(fill="x", padx=3, pady=3)
        
        # 선택 상태에 따른 색상 설정
        is_selected = index in self.selected_indices
        if is_selected:
            item_frame.configure(fg_color=("lightblue", "darkblue"))
        
        # 체크박스
        checkbox_var = tk.BooleanVar(value=is_selected)
        checkbox = ctk.CTkCheckBox(item_frame, text="", variable=checkbox_var,
                                 command=lambda: self.toggle_selection(index),
                                 width=18, height=18)
        checkbox.pack(side="left", padx=(8, 3), pady=8)
        
        # 순서 번호
        order_label = ctk.CTkLabel(item_frame, text=f"{index + 1}.", 
                                 font=ctk.CTkFont(size=13, weight="bold"))
        order_label.pack(side="left", padx=(3, 3), pady=8)
        
        # 제목
        title_label = ctk.CTkLabel(item_frame, text=praise['title'], 
                                 font=ctk.CTkFont(size=13, weight="bold"),
                                 anchor="w")
        title_label.pack(side="left", fill="x", expand=True, padx=(0, 8), pady=8)
        
        # 버튼 프레임 (오른쪽)
        button_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        button_frame.pack(side="right", padx=(0, 8), pady=3)
        
        # 목록에서 빼기 버튼 (X 버튼)
        remove_button = ctk.CTkButton(button_frame, text="×", 
                                    command=lambda: self.remove_selected_by_index(index),
                                    width=28, height=28, font=ctk.CTkFont(size=14, weight="bold"),
                                    fg_color="orange", hover_color="darkorange")
        remove_button.pack(side="right", padx=(0, 3))
        
        # 파일 삭제 버튼 (휴지통 아이콘)
        delete_button = ctk.CTkButton(button_frame, text="🗑", 
                                    command=lambda: self.delete_pptx_file(praise),
                                    width=28, height=28, font=ctk.CTkFont(size=12),
                                    fg_color="red", hover_color="darkred")
        delete_button.pack(side="right", padx=(0, 3))
        
        # 툴팁 추가
        self.create_tooltip(remove_button, "목록에서 빼기 (파일은 유지)")
        self.create_tooltip(delete_button, "파일 완전 삭제 (되돌릴 수 없음)")
        
        # 드래그-앤-드롭: 순서 변경
        def on_drag_start(event):
            self.dragging_index = index
            self.drag_start_y = event.y_root
            try:
                item_frame.configure(fg_color=("lightgray", "gray"))
            except Exception:
                pass
        
        def on_drag_motion(event):
            if self.dragging_index is not None:
                self._update_drop_indicator(event.y_root)
        
        def on_drag_end(event):
            try:
                # 드랍 위치 계산 (마우스 y 위치와 각 항목의 중앙 y 비교)
                children = self.selected_frame.winfo_children()
                drop_index = len(children) - 1
                mouse_y_root = event.y_root
                for i, child in enumerate(children):
                    center_y = child.winfo_rooty() + (child.winfo_height() // 2)
                    if mouse_y_root <= center_y:
                        drop_index = i
                        break
                
                if self.dragging_index is not None and drop_index is not None and self.dragging_index != drop_index:
                    self._reorder_selected(self.dragging_index, drop_index)
            finally:
                self.dragging_index = None
                self.drag_start_y = None
                self._clear_drop_indicator()
                # 색상 복원
                is_selected_local = index in self.selected_indices
                try:
                    if is_selected_local:
                        item_frame.configure(fg_color=("lightblue", "darkblue"))
                    else:
                        item_frame.configure(fg_color=("gray90", "gray20"))
                except Exception:
                    pass
        
        # 바인딩: 체크박스와 버튼 영역은 제외, 아이템 프레임과 텍스트에만 바인딩
        item_frame.bind("<ButtonPress-1>", on_drag_start)
        item_frame.bind("<B1-Motion>", on_drag_motion)
        item_frame.bind("<ButtonRelease-1>", on_drag_end)
        title_label.bind("<ButtonPress-1>", on_drag_start)
        title_label.bind("<B1-Motion>", on_drag_motion)
        title_label.bind("<ButtonRelease-1>", on_drag_end)
        order_label.bind("<ButtonPress-1>", on_drag_start)
        order_label.bind("<B1-Motion>", on_drag_motion)
        order_label.bind("<ButtonRelease-1>", on_drag_end)
        
        # 체크박스 상태 업데이트를 위한 참조 저장
        item_frame.checkbox = checkbox
        item_frame.checkbox_var = checkbox_var

    def _update_drop_indicator(self, mouse_y_root):
        """드롭 위치 표시선 업데이트"""
        try:
            # 기존 표시선 제거
            self._clear_drop_indicator()
            
            # 스크롤 영역 내에서만 표시
            scroll_y = self.selected_frame.winfo_rooty()
            scroll_height = self.selected_frame.winfo_height()
            
            if not (scroll_y <= mouse_y_root <= scroll_y + scroll_height):
                return
            
            # 드롭 위치 계산
            children = self.selected_frame.winfo_children()
            if not children:
                return
            
            drop_y = scroll_y + 10  # 기본값 (맨 위)
            
            for i, child in enumerate(children):
                child_y = child.winfo_rooty()
                child_height = child.winfo_height()
                child_center = child_y + (child_height // 2)
                
                if mouse_y_root <= child_center:
                    drop_y = child_y
                    break
                else:
                    drop_y = child_y + child_height
            
            # 표시선 생성
            self.drop_indicator = ctk.CTkFrame(
                self.selected_frame,
                height=3,
                fg_color=("blue", "lightblue"),
                corner_radius=0
            )
            self.drop_indicator.place(x=10, y=drop_y - scroll_y, relwidth=0.95)
            
        except Exception as e:
            print(f"[DEBUG] 드롭 표시선 업데이트 실패: {e}")
    
    def _clear_drop_indicator(self):
        """드롭 위치 표시선 제거"""
        try:
            if self.drop_indicator:
                self.drop_indicator.destroy()
                self.drop_indicator = None
        except Exception:
            pass
    
    def _reorder_selected(self, from_index, to_index):
        """선택된 찬양 목록에서 from_index 항목을 to_index 위치로 이동"""
        if not (0 <= from_index < len(self.selected_praises)):
            return
        if not (0 <= to_index < len(self.selected_praises)):
            to_index = len(self.selected_praises) - 1
        if from_index == to_index:
            return
        
        item = self.selected_praises.pop(from_index)
        self.selected_praises.insert(to_index, item)
        
        # 선택 인덱스 재매핑
        remapped = set()
        for idx in self.selected_indices:
            if idx == from_index:
                remapped.add(to_index)
            elif from_index < to_index and from_index < idx <= to_index:
                remapped.add(idx - 1)
            elif to_index < from_index and to_index <= idx < from_index:
                remapped.add(idx + 1)
            else:
                remapped.add(idx)
        self.selected_indices = remapped
        self.update_selected_display()
    
    def toggle_selection(self, index):
        """항목 선택 토글"""
        if index in self.selected_indices:
            self.selected_indices.remove(index)
        else:
            self.selected_indices.add(index)
        self.update_selected_display()
    
    def update_checkbox_states(self):
        """체크박스 상태 업데이트"""
        for i, widget in enumerate(self.selected_frame.winfo_children()):
            if hasattr(widget, 'checkbox_var'):
                is_selected = i in self.selected_indices
                widget.checkbox_var.set(is_selected)
    
    
    def delete_selected_items(self):
        """선택된 항목들 삭제"""
        if not self.selected_indices:
            messagebox.showwarning("경고", "삭제할 항목을 선택해주세요.")
            return
        
        # 인덱스를 내림차순으로 정렬하여 뒤에서부터 삭제
        sorted_indices = sorted(self.selected_indices, reverse=True)
        
        for index in sorted_indices:
            if 0 <= index < len(self.selected_praises):
                del self.selected_praises[index]
        
        # 선택 상태 초기화
        self.selected_indices.clear()
        self.update_selected_display()
        messagebox.showinfo("완료", f"{len(sorted_indices)}개 항목이 삭제되었습니다.")
    
    def remove_selected_by_index(self, index):
        """인덱스로 선택된 찬양 제거"""
        if 0 <= index < len(self.selected_praises):
            del self.selected_praises[index]
            # 선택된 인덱스들 업데이트
            new_selected_indices = set()
            for idx in self.selected_indices:
                if idx > index:
                    new_selected_indices.add(idx - 1)
                elif idx < index:
                    new_selected_indices.add(idx)
            self.selected_indices = new_selected_indices
            self.update_selected_display()
    
    def select_all_items(self):
        """전체 선택/해제 토글"""
        if len(self.selected_indices) == len(self.selected_praises):
            # 모두 선택된 상태면 모두 해제
            self.selected_indices.clear()
            self.select_all_button.configure(text="전체 선택")
        else:
            # 일부만 선택되거나 아무것도 선택되지 않은 상태면 모두 선택
            self.selected_indices = set(range(len(self.selected_praises)))
            self.select_all_button.configure(text="전체 해제")
        
            self.update_selected_display()
    
    
    
    def clear_selected(self):
        """선택된 찬양 전체 제거"""
        self.selected_praises.clear()
        self.selected_indices.clear()
        self.select_all_button.configure(text="전체 선택")
        self.update_selected_display()
    
    def generate_ppt(self):
        """PPT 생성"""
        if not self.selected_praises:
            messagebox.showwarning("경고", "선택된 찬양이 없습니다.")
            return
        
        if not self.generator:
            messagebox.showerror("오류", "PPT 생성기가 초기화되지 않았습니다.")
            return
        
        # 파일 저장 대화상자
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx")],
            title="PPT 저장 위치 선택"
        )
        
        if not output_path:
            return
        
        def generate_thread():
            try:
                self.progress_var.set("PPT 생성 중...")
                self.root.update()
                
                result = self.generator.create_ppt_from_lyrics(self.selected_praises, output_path)
                
                if result:
                    self.progress_var.set("PPT 생성 완료")
                    # 실제 저장된 파일 경로 확인
                    import os
                    if os.path.exists(output_path):
                        messagebox.showinfo("완료", f"PPT가 생성되었습니다:\n{output_path}")
                    else:
                        # 대체 파일명으로 저장된 경우 찾기
                        import glob
                        base_name = os.path.splitext(output_path)[0]
                        pattern = f"{base_name}_*.pptx"
                        alt_files = glob.glob(pattern)
                        if alt_files:
                            alt_file = alt_files[0]  # 가장 최근 파일
                            messagebox.showinfo("완료", f"PPT가 생성되었습니다 (대체 파일명):\n{alt_file}")
                        else:
                            messagebox.showinfo("완료", "PPT가 생성되었습니다.")
                else:
                    self.progress_var.set("PPT 생성 실패")
                    messagebox.showerror("오류", "PPT 생성에 실패했습니다.\n파일이 다른 프로그램에서 사용 중일 수 있습니다.")
            except Exception as e:
                self.progress_var.set("PPT 생성 실패")
                messagebox.showerror("오류", f"PPT 생성 실패: {e}")
        
        threading.Thread(target=generate_thread, daemon=True).start()
    
    def delete_pptx_file(self, praise):
        """PPTX 파일 삭제"""
        try:
            # 확인 대화상자
            result = messagebox.askyesno("파일 삭제", 
                                       f"'{praise['title']}' 파일을 삭제하시겠습니까?\n\n"
                                       f"파일: {praise['filename']}\n"
                                       f"이 작업은 되돌릴 수 없습니다.")
            
            if not result:
                return
            
            # 파일 경로 확인 (여러 경로 시도)
            file_path_str = praise['file_path']
            print(f"[DEBUG] 원본 경로: {file_path_str}")
            
            # 경로 수정
            if file_path_str.startswith("praise_indexer\\"):
                file_path_str = file_path_str.replace("praise_indexer\\", "")
            elif file_path_str.startswith("praise_indexer/"):
                file_path_str = file_path_str.replace("praise_indexer/", "")
            
            print(f"[DEBUG] 수정된 경로: {file_path_str}")
            
            # 여러 경로 시도
            possible_paths = [
                Path(file_path_str),  # 상대 경로
                Path.cwd() / file_path_str,  # 현재 디렉토리 기준
                Path("praise_indexer") / file_path_str,  # praise_indexer 디렉토리 기준
            ]
            
            file_path = None
            for path in possible_paths:
                print(f"[DEBUG] 시도 중: {path}")
                if path.exists():
                    file_path = path
                    print(f"[DEBUG] 파일 발견: {path}")
                    break
            
            if not file_path:
                # 파일이 존재하지 않는 경우, JSON에서만 제거
                messagebox.showwarning("경고", 
                                     f"실제 파일이 존재하지 않습니다.\n"
                                     f"시도한 경로들:\n" + 
                                     "\n".join([str(p) for p in possible_paths]) +
                                     f"\n\nJSON에서만 데이터를 제거합니다.")
                
                # JSON에서 해당 항목 제거
                self.indexer.remove_praise_by_id(praise['id'])
                
                # JSON 파일 저장
                self.indexer.save_to_json()
                
                # UI에서 해당 항목만 제거 (선택 목록과 검색 결과 유지)
                try:
                    removed_id = praise['id']
                    # 선택 목록에서 제거
                    self.selected_praises = [p for p in self.selected_praises if p.get('id') != removed_id]
                    # 선택 인덱스 재계산
                    self.selected_indices = {i for i in self.selected_indices if i < len(self.selected_praises)}
                    self.update_selected_display()
                    # 검색 결과에서도 제거
                    self.search_results = [p for p in self.search_results if p.get('id') != removed_id]
                    self.update_results_display()
                except Exception:
                    pass
                
                messagebox.showinfo("완료", f"'{praise['title']}' 데이터가 JSON에서 제거되었습니다.")
                return
            
            # 파일 삭제 (재시도 로직 포함)
            import time
            max_retries = 3
            
            for attempt in range(max_retries):
                try:
                    file_path.unlink()
                    print(f"[OK] 파일 삭제됨: {file_path}")
                    break
                except PermissionError as e:
                    if attempt < max_retries - 1:
                        print(f"[WARNING] 파일 삭제 실패 (시도 {attempt + 1}/{max_retries}): {e}")
                        time.sleep(1)  # 1초 대기 후 재시도
                    else:
                        print(f"[ERROR] 파일 삭제 최종 실패: {file_path}")
                        messagebox.showwarning("경고", 
                                             f"파일을 삭제할 수 없습니다:\n{file_path}\n\n"
                                             f"파일이 다른 프로그램에서 사용 중일 수 있습니다.\n"
                                             f"JSON에서만 데이터를 제거합니다.")
                        # JSON에서만 제거하고 계속 진행
                        break
                except Exception as e:
                    print(f"[ERROR] 파일 삭제 실패: {e}")
                    messagebox.showwarning("경고", f"파일 삭제 실패: {e}\nJSON에서만 데이터를 제거합니다.")
                    break
            
            # JSON에서 해당 항목 제거
            self.indexer.remove_praise_by_id(praise['id'])
            
            # JSON 파일 저장
            self.indexer.save_to_json()
            
            # UI에서 해당 항목만 제거 (선택 목록과 검색 결과 유지)
            try:
                removed_id = praise['id']
                # 선택 목록에서 제거
                self.selected_praises = [p for p in self.selected_praises if p.get('id') != removed_id]
                # 선택 인덱스 재계산
                self.selected_indices = {i for i in self.selected_indices if i < len(self.selected_praises)}
                self.update_selected_display()
                # 검색 결과에서도 제거
                self.search_results = [p for p in self.search_results if p.get('id') != removed_id]
                self.update_results_display()
            except Exception:
                pass
            
            messagebox.showinfo("완료", f"'{praise['title']}' 파일이 삭제되었습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"파일 삭제 실패: {e}")
            print(f"[ERROR] 파일 삭제 실패: {e}")
    
    def add_pptx_file(self):
        """새 PPTX 파일 추가 (복사 없이 직접 인덱싱)"""
        try:
            # 현재 UI 상태 백업 (선택 목록 + 검색 상태)
            backup_selected = list(self.selected_praises)
            backup_indices = set(self.selected_indices)
            backup_query = self.search_var.get()
            backup_search_type = self.search_type_var.get()

            # 파일 선택 대화상자
            file_paths = filedialog.askopenfilenames(
                title="추가할 PPTX 파일 선택",
                filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")]
            )
            
            if not file_paths:
                return
            
            self.progress_var.set("새 파일 인덱싱 중...")
            self.root.update()

            added_count = 0
            for file_path in file_paths:
                ok = self.indexer.add_single_file(file_path)
                if ok:
                    added_count += 1
            
            # JSON 파일 저장
            self.indexer.save_to_json()
            
            # 인덱서 최신 데이터 반영 (필요 시)
            self.indexer.load_from_json()
            
            # 검색 상태 복원 및 재검색 (결과 유지)
            try:
                if backup_query:
                    self.search_var.set(backup_query)
                    self.search_type_var.set(backup_search_type)
                    self.perform_search()
                else:
                    # 검색어가 없던 경우에도 결과 영역 유지
                    self.update_results_display()
            except Exception:
                self.update_results_display()
            
            # 선택 목록 복원
            self.selected_praises = backup_selected
            self.selected_indices = backup_indices
            self.update_selected_display()
            
            self.progress_var.set("파일 추가 완료")
            messagebox.showinfo("완료", f"{added_count}개 파일이 추가되었습니다.")
            
        except Exception as e:
            self.progress_var.set("파일 추가 실패")
            messagebox.showerror("오류", f"파일 추가 실패: {e}")
            print(f"[ERROR] 파일 추가 실패: {e}")
    
    def refresh_data(self):
        """데이터 새로고침"""
        try:
            # 데이터 다시 로드
            self.load_data()
            
            # 검색 결과 초기화
            self.search_results = []
            self.update_results_display()
            
            # 선택된 찬양 초기화
            self.selected_praises = []
            self.update_selected_display()
            
            print("[OK] 데이터 새로고침 완료")
            
        except Exception as e:
            print(f"[ERROR] 데이터 새로고침 실패: {e}")
    
    def run(self):
        """GUI 실행"""
        self.root.mainloop()

def main():
    """메인 함수"""
    app = JSONPraiseGUI()
    app.run()

if __name__ == "__main__":
    main()