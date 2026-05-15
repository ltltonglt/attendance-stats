from __future__ import annotations

import os
import sys
import threading
from pathlib import Path


def configure_tcl_tk_library() -> None:
    python_root = Path(sys.base_prefix)
    tcl_library = python_root / "tcl" / "tcl8.6"
    tk_library = python_root / "tcl" / "tk8.6"
    if "TCL_LIBRARY" not in os.environ and tcl_library.exists():
        os.environ["TCL_LIBRARY"] = str(tcl_library)
    if "TK_LIBRARY" not in os.environ and tk_library.exists():
        os.environ["TK_LIBRARY"] = str(tk_library)


configure_tcl_tk_library()

from tkinter import END, BooleanVar, Button, Checkbutton, Entry, Label, StringVar, Tk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from attendance_audit_core import AttendanceAuditError, audit_personnel


class AttendanceAuditApp(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("人员打卡核查工具")
        self.geometry("780x420")
        self.minsize(700, 360)

        cwd = Path.cwd()
        default_attendance_dir = cwd / "打卡数据导出" / "20260515-142723"
        if not default_attendance_dir.exists():
            default_attendance_dir = cwd / "打卡数据导出"
        self.person_file_var = StringVar(value=str(cwd / "人员信息.xlsx"))
        self.attendance_dir_var = StringVar(value=str(default_attendance_dir))
        self.output_file_var = StringVar(value=str(cwd / "人员打卡核查结果.xlsx"))
        self.merge_attendance_var = BooleanVar(value=False)
        self.status_var = StringVar(value="请选择文件后开始核查")

        self._build_layout()

    def _build_layout(self) -> None:
        self.columnconfigure(1, weight=1)
        self.rowconfigure(5, weight=1)

        Label(self, text="人员信息表").grid(row=0, column=0, padx=12, pady=(14, 6), sticky="w")
        Entry(self, textvariable=self.person_file_var).grid(row=0, column=1, padx=8, pady=(14, 6), sticky="ew")
        Button(self, text="选择文件", command=self.choose_person_file).grid(row=0, column=2, padx=12, pady=(14, 6))

        Label(self, text="打卡数据文件夹").grid(row=1, column=0, padx=12, pady=6, sticky="w")
        Entry(self, textvariable=self.attendance_dir_var).grid(row=1, column=1, padx=8, pady=6, sticky="ew")
        Button(self, text="选择文件夹", command=self.choose_attendance_dir).grid(row=1, column=2, padx=12, pady=6)

        Label(self, text="输出文件").grid(row=2, column=0, padx=12, pady=6, sticky="w")
        Entry(self, textvariable=self.output_file_var).grid(row=2, column=1, padx=8, pady=6, sticky="ew")
        Button(self, text="保存为", command=self.choose_output_file).grid(row=2, column=2, padx=12, pady=6)

        Checkbutton(
            self,
            text="同时合并打卡数据",
            variable=self.merge_attendance_var,
            anchor="w",
        ).grid(row=3, column=0, columnspan=3, padx=12, pady=(8, 2), sticky="w")

        self.run_button = Button(self, text="开始核查", command=self.start_audit, height=2)
        self.run_button.grid(row=4, column=0, columnspan=3, padx=12, pady=12, sticky="ew")

        self.log_text = ScrolledText(self, height=10, state="disabled")
        self.log_text.grid(row=5, column=0, columnspan=3, padx=12, pady=6, sticky="nsew")

        Label(self, textvariable=self.status_var, anchor="w").grid(
            row=6, column=0, columnspan=3, padx=12, pady=(4, 12), sticky="ew"
        )

    def choose_person_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择人员信息表",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.person_file_var.set(path)
            default_output = Path(path).with_name("人员打卡核查结果.xlsx")
            self.output_file_var.set(str(default_output))

    def choose_attendance_dir(self) -> None:
        path = filedialog.askdirectory(title="选择打卡数据导出文件夹")
        if path:
            self.attendance_dir_var.set(path)

    def choose_output_file(self) -> None:
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile=Path(self.output_file_var.get()).name,
        )
        if path:
            self.output_file_var.set(path)

    def start_audit(self) -> None:
        person_file = self.person_file_var.get().strip()
        attendance_dir = self.attendance_dir_var.get().strip()
        output_file = self.output_file_var.get().strip()
        merge_attendance = self.merge_attendance_var.get()

        if not person_file or not attendance_dir or not output_file:
            messagebox.showwarning("缺少路径", "请先选择人员信息表、打卡数据文件夹和输出文件。")
            return
        if Path(person_file).resolve() == Path(output_file).resolve():
            messagebox.showwarning("输出文件错误", "输出文件不能与原始人员信息表相同。")
            return

        self.run_button.config(state="disabled")
        self.clear_log()
        self.append_log("开始核查...")
        self.status_var.set("正在处理")

        thread = threading.Thread(
            target=self._run_audit_worker,
            args=(person_file, attendance_dir, output_file, merge_attendance),
            daemon=True,
        )
        thread.start()

    def _run_audit_worker(
        self,
        person_file: str,
        attendance_dir: str,
        output_file: str,
        merge_attendance: bool,
    ) -> None:
        try:
            merged_output_file = None
            if merge_attendance:
                output_path = Path(output_file)
                merged_output_file = output_path.with_name(f"{output_path.stem}_合并打卡数据.xlsx")
            summary = audit_personnel(
                person_file,
                attendance_dir,
                output_file,
                progress_callback=lambda message: self.after(0, self.append_log, message),
                merge_attendance=merge_attendance,
                merged_output_file=merged_output_file,
            )
        except (AttendanceAuditError, OSError, ValueError) as exc:
            self.after(0, self.audit_failed, str(exc))
            return
        except Exception as exc:
            self.after(0, self.audit_failed, f"未知错误：{exc}")
            return
        self.after(0, self.audit_succeeded, summary)

    def audit_succeeded(self, summary: object) -> None:
        self.run_button.config(state="normal")
        self.status_var.set("核查完成")
        duplicate_names = ", ".join(summary.duplicate_names) if summary.duplicate_names else "无"
        self.append_log(
            f"人员数量：{summary.person_count}\n"
            f"打卡文件：{summary.attendance_file_count}\n"
            f"打卡记录：{summary.attendance_record_count}\n"
            f"重复姓名：{duplicate_names}\n"
            f"输出文件：{summary.output_path}"
        )
        if summary.merged_output_path:
            self.append_log(f"合并打卡数据：{summary.merged_output_path}")
        messagebox.showinfo("完成", f"核查完成，结果已保存到：\n{summary.output_path}")

    def audit_failed(self, message: str) -> None:
        self.run_button.config(state="normal")
        self.status_var.set("核查失败")
        self.append_log(f"错误：{message}")
        messagebox.showerror("核查失败", message)

    def clear_log(self) -> None:
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", END)
        self.log_text.config(state="disabled")

    def append_log(self, message: str) -> None:
        self.log_text.config(state="normal")
        self.log_text.insert(END, message + "\n")
        self.log_text.see(END)
        self.log_text.config(state="disabled")
        self.status_var.set(message)


def main() -> None:
    app = AttendanceAuditApp()
    app.mainloop()


if __name__ == "__main__":
    main()
