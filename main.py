import sys
import subprocess
import platform
import ctypes
import os
import json
import customtkinter as ctk



# Translation dictionary
TRANSLATIONS = {
    "ko": {
        "title": "네트워크 어댑터 우선순위 관리자",
        "header": "연결된 네트워크 어댑터",
        "description": "드래그 앤 드롭으로 우선순위를 변경하세요",
        "refresh": "새로고침",
        "language": "English",
        "adapter_type_unknown": "알 수 없음",
        "adapter_type_wireless": "무선",
        "adapter_type_wired": "유선",
        "status_no_header": "헤더를 찾을 수 없습니다.",
        "status_error": "오류 발생: {}",
        "status_adapters_found": "총 {}개의 연결된 어댑터를 찾았습니다.",
        "status_no_adapters": "연결된 네트워크 어댑터가 없습니다.",
        "status_priority_changed": "우선순위가 성공적으로 변경되었습니다.",
        "status_priority_failed": "우선순위 변경 실패: {}"
    },
    "en": {
        "title": "Network Adapter Priority Manager",
        "header": "Connected Network Adapters",
        "description": "Drag and drop to change priority",
        "refresh": "Refresh",
        "language": "한국어",
        "adapter_type_unknown": "Unknown",
        "adapter_type_wireless": "Wireless",
        "adapter_type_wired": "Wired",
        "status_no_header": "Could not find header.",
        "status_error": "Error occurred: {}",
        "status_adapters_found": "Found {} connected adapter(s).",
        "status_no_adapters": "No connected network adapters found.",
        "status_priority_changed": "Priority successfully changed.",
        "status_priority_failed": "Failed to change priority: {}"
    }
}

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if is_admin():
        return True
    else:
        ctypes.windll.shell32.ShellExecuteW(
            None, 
            "runas", 
            sys.executable, 
            " ".join([os.path.abspath(sys.argv[0])] + sys.argv[1:]),
            None, 
            1
        )
        return False

class DraggableListBox(ctk.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.items = []
        self.selected_index = None
        self.drag_start_y = None
        self.root = None  # 루트 윈도우 참조 저장용
        
    def set_root(self, root):
        self.root = root
        
    def add_item(self, text, data=None):
        item = ctk.CTkFrame(self, fg_color="transparent")
        item.pack(fill="x", padx=5, pady=2)
        
        label = ctk.CTkLabel(item, text=text, anchor="w", padx=10)
        label.pack(fill="x", expand=True)
        
        # 드래그 이벤트 바인딩
        label.bind("<Button-1>", lambda e, i=len(self.items): self.start_drag(e, i))
        label.bind("<B1-Motion>", self.drag)
        label.bind("<ButtonRelease-1>", self.end_drag)
        
        self.items.append({"frame": item, "label": label, "data": data})
        
    def clear(self):
        for item in self.items:
            item["frame"].destroy()
        self.items.clear()
        
    def start_drag(self, event, index):
        self.selected_index = index
        self.drag_start_y = event.y_root
        self.items[index]["frame"].configure(fg_color=("gray85", "gray25"))
        
    def drag(self, event):
        if self.selected_index is None:
            return
            
        delta_y = event.y_root - self.drag_start_y
        new_index = self.selected_index
        
        # 위치 계산
        item_height = self.items[0]["frame"].winfo_height()
        move_items = int(delta_y / item_height)
        
        if move_items != 0:
            new_index = max(0, min(len(self.items)-1, self.selected_index + move_items))
            if new_index != self.selected_index:
                # 아이템 순서 변경
                item = self.items.pop(self.selected_index)
                self.items.insert(new_index, item)
                self.repack_items()
                self.selected_index = new_index
                self.drag_start_y = event.y_root
                # 메트릭 업데이트
                if self.root:
                    self.root.update_metrics()
                
    def end_drag(self, event):
        if self.selected_index is not None:
            self.items[self.selected_index]["frame"].configure(fg_color="transparent")
        self.selected_index = None
        
    def repack_items(self):
        for item in self.items:
            item["frame"].pack_forget()
        for item in self.items:
            item["frame"].pack(fill="x", padx=5, pady=2)

class NetworkAdapterPriorityManager(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Initialize language
        self.current_lang = "en"
        self.load_language_preference()
        
        # Theme settings
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        
        self.title(TRANSLATIONS[self.current_lang]["title"])
        self.geometry("600x500")
        
        # Main frame
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        main_frame = ctk.CTkFrame(self)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_rowconfigure(2, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Title
        self.title_label = ctk.CTkLabel(main_frame, 
                           text=TRANSLATIONS[self.current_lang]["header"], 
                           font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.grid(row=0, column=0, pady=(0, 10))
        
        # Description
        self.description_label = ctk.CTkLabel(main_frame, 
                                text=TRANSLATIONS[self.current_lang]["description"],
                                font=ctk.CTkFont(size=14))
        self.description_label.grid(row=1, column=0, pady=(0, 10))
        
        # Adapter list
        self.adapter_list = DraggableListBox(main_frame, width=500, height=300)
        self.adapter_list.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        self.adapter_list.set_root(self)
        
        # Button frame
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.grid(row=3, column=0, pady=(0, 10))
        
        # Language button
        self.lang_btn = ctk.CTkButton(button_frame,
                                   text=TRANSLATIONS[self.current_lang]["language"],
                                   command=self.toggle_language)
        self.lang_btn.pack(side="left", padx=5)
        
        # Refresh button
        self.refresh_btn = ctk.CTkButton(button_frame, 
                                  text=TRANSLATIONS[self.current_lang]["refresh"],
                                  command=self.refresh_adapters)
        self.refresh_btn.pack(side="left", padx=5)
        
        # Status label
        self.status_label = ctk.CTkLabel(main_frame, text="")
        self.status_label.grid(row=4, column=0)
        
        # Load initial adapter list
        self.refresh_adapters()
        
    def load_language_preference(self):
        self.current_lang = "en"
            
    def save_language_preference(self):
        pass
        try:
            if os.path.exists("language.json"):
                with open("language.json", "r") as f:
                    data = json.load(f)
                    self.current_lang = data.get("language", "ko")
        except:
            self.current_lang = "ko"
            
    def toggle_language(self):
        self.current_lang = "en" if self.current_lang == "ko" else "ko"
        self.save_language_preference()
        
        # Update UI text
        self.title(TRANSLATIONS[self.current_lang]["title"])
        self.title_label.configure(text=TRANSLATIONS[self.current_lang]["header"])
        self.description_label.configure(text=TRANSLATIONS[self.current_lang]["description"])
        self.refresh_btn.configure(text=TRANSLATIONS[self.current_lang]["refresh"])
        self.lang_btn.configure(text=TRANSLATIONS[self.current_lang]["language"])
        
        # Refresh adapter list to update adapter type text
        self.refresh_adapters()
        
        # 초기 어댑터 목록 로드
        self.refresh_adapters()
        
    def get_adapters(self):
        adapters = []
        try:
            output_bytes = subprocess.check_output(
                ["netsh", "interface", "ipv4", "show", "interfaces"]
            )
            
            encodings = ['utf-8', 'cp949', 'euc-kr']
            output = None
            
            for encoding in encodings:
                try:
                    output = output_bytes.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue
            
            if output is None:
                output = output_bytes.decode('utf-8', errors='replace')
            
            lines = output.split('\n')
            # 헤더 찾기
            header_index = -1
            for i, line in enumerate(lines):
                if "---" in line:
                    header_index = i
                    break
            
            if header_index == -1:
                raise Exception(TRANSLATIONS[self.current_lang]["status_no_header"])
            
            # 헤더 다음 줄부터 처리
            for line in lines[header_index + 1:]:
                if not line.strip():
                    continue
                
                try:
                    # 공백으로 분리
                    parts = [p for p in line.split() if p]  # 빈 문자열 제거
                    if len(parts) >= 4:
                        idx = parts[0]
                        metric = parts[1]
                        state = parts[3]  # MTU는 건너뛰고 상태 가져오기
                        name = ' '.join(parts[4:])  # 나머지는 이름
                        
                        # 연결된 어댑터만 추가
                        if "connected" in state.lower():
                            adapter_type = TRANSLATIONS[self.current_lang]["adapter_type_unknown"]
                            name_lower = name.lower()
                            if any(x in name_lower for x in ["wi-fi", "wireless", "무선"]):
                                adapter_type = TRANSLATIONS[self.current_lang]["adapter_type_wireless"]
                            elif any(x in name_lower for x in ["ethernet", "이더넷"]):
                                adapter_type = TRANSLATIONS[self.current_lang]["adapter_type_wired"]
                            elif "loopback" in name_lower or "vpn" in name_lower:
                                continue  # 루프백 어댑터와 vpn 어댑터 제외
                            
                            try:
                                metric_value = int(metric)
                                adapters.append({
                                    "name": name,
                                    "type": adapter_type,
                                    "metric": metric_value
                                })
                            except ValueError:
                                continue
                            
                except (ValueError, IndexError) as e:
                    continue
            
            # 메트릭 값으로 정렬
            adapters.sort(key=lambda x: x['metric'])
            
        except Exception as e:
            error_msg = TRANSLATIONS[self.current_lang]["status_error"].format(str(e))
            self.status_label.configure(text=error_msg)
        
        return adapters

    def refresh_adapters(self):
        self.adapter_list.clear()
        adapters = self.get_adapters()
        
        if adapters:
            for adapter in adapters:
                self.adapter_list.add_item(
                    f"{adapter['name']} ({adapter['type']}) - 메트릭: {adapter['metric']}",
                    adapter
                )
            self.status_label.configure(text=TRANSLATIONS[self.current_lang]["status_adapters_found"].format(len(adapters)))
        else:
            self.status_label.configure(text=TRANSLATIONS[self.current_lang]["status_no_adapters"])

    def update_metrics(self):
        try:
            # Create a batch command string
            commands = []
            base_metric = 1
            for i, item in enumerate(self.adapter_list.items):
                adapter = item["data"]
                new_metric = base_metric + (i * 10)
                commands.append(f'netsh interface ipv4 set interface "{adapter["name"]}" metric={new_metric}')
            
            # Execute all commands in a single subprocess call
            batch_command = ' & '.join(commands)
            subprocess.run(batch_command, shell=True, check=True)
            
            self.refresh_adapters()
            self.status_label.configure(text=TRANSLATIONS[self.current_lang]["status_priority_changed"])
        except Exception as e:
            error_msg = TRANSLATIONS[self.current_lang]["status_priority_failed"].format(str(e))
            self.status_label.configure(text=error_msg)

if __name__ == "__main__":
    if platform.system() == "Windows":
        if not run_as_admin():
            sys.exit()
    
    app = NetworkAdapterPriorityManager()
    app.mainloop()
