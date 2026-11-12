"""印章候选仅存于内存，确认后才交给 Word 导出。"""

from io import BytesIO
from dataclasses import replace
import math
from time import perf_counter
from pathlib import Path
from queue import Empty, Queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import Image, ImageTk

from .extraction import candidate_from_bytes, extract_docx_images
from .pdf_pages import PdfPages, PdfPageError, file_signature


class SealPanel(ctk.CTkFrame):
    def __init__(self, master, current_file, **kwargs):
        super().__init__(master, **kwargs)
        self.current_file = current_file
        self.file = None
        self.candidates = []
        self.thumbnail_cache = {}  # 只缓存小尺寸预览，排序和勾选无需反复解码原图。
        self.selected = []
        self.generation = 0
        self.windows = []
        self.pdf_reader = PdfPages()
        self.pdf_source = None
        self.pdf_loading = False
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        ctk.CTkLabel(self, text="电子印章采集").grid(row=0, column=0, sticky="w", padx=12)
        controls = ctk.CTkFrame(self, fg_color="transparent")
        controls.grid(row=1, column=0, sticky="ew", padx=8)
        self.source = tk.StringVar(value="当前申请表 PDF")  # 来源只负责选择文件，候选流程共用。
        self.source_label_var = tk.StringVar()
        ctk.CTkButton(controls, text="当前 PDF", width=90, command=self.use_current_pdf).pack(side="left", padx=4)
        ctk.CTkButton(controls, text="选择其他文件", width=120, command=self.choose_file).pack(side="left", padx=4)
        ctk.CTkButton(controls, text="提取印章", width=90, command=self.extract).pack(side="left", padx=4)
        ctk.CTkLabel(controls, textvariable=self.source_label_var, anchor="w", justify="left",
                     wraplength=400).pack(side="left", fill="x", expand=True, padx=6)
        self.status = ctk.CTkLabel(self, text="尚未选择印章", anchor="w", wraplength=760, justify="left")
        self.status.grid(row=2, column=0, sticky="ew", padx=12)
        self.cards = ctk.CTkScrollableFrame(self, height=150)
        self.cards.grid(row=3, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.winfo_toplevel().bind("<MouseWheel>", self._scroll_candidates, add="+")
        self._refresh_source_label()

    def _refresh_source_label(self):
        if self.source.get() == "当前申请表 PDF":
            path = self.current_file()
            text = f"当前 PDF：{path.name}" if path is not None and path.suffix.lower() == ".pdf" else "当前 PDF：上方尚未选择 PDF"
        else:
            text = f"其他文件：{self.file.name}" if self.file is not None else "其他文件：尚未选择"
        self.source_label_var.set(text)

    def _scroll_candidates(self, event):
        # 在顶层绑定阶段消费事件，避免继续触发 CTk 的全局滚轮绑定。
        if event.delta and self.cards.check_if_master_is_canvas(event.widget):
            self.cards._parent_canvas.yview_scroll(-int(event.delta / 6) or (-1 if event.delta > 0 else 1), "units")
            return "break"

    def clear(self):
        self.generation += 1
        self.candidates = []
        self.thumbnail_cache.clear()
        self.selected = []
        self.pdf_source = None
        self.pdf_loading = False
        for window in list(self.windows):
            if window.winfo_exists():
                window.destroy()
        self.windows = []
        for widget in self.cards.winfo_children():
            widget.destroy()
        self.status.configure(text="尚未选择印章；生成 Word 时采集框留空")
        self._refresh_source_label()

    def reset_customer(self):
        self.use_current_pdf()

    def pdf_progress(self, token, text):
        if token == self.generation:
            self.status.configure(text=text)

    def accept_pdf_source(self, token, source):
        if token == self.generation and self.source.get() == "当前申请表 PDF":
            self._install_pdf_source(source)
            self.status.configure(text=source.notice)
            self.render()

    def use_current_pdf(self):
        self.source.set("当前申请表 PDF")
        self.file = None
        self.clear()

    def choose_file(self):
        selected = filedialog.askopenfilename(title="选择印章资料", filetypes=[("印章资料", "*.docx *.pdf *.png *.jpg *.jpeg")])
        if selected:
            path = Path(selected)
            if path.suffix.lower() not in {".pdf", ".docx", ".png", ".jpg", ".jpeg"} or not path.is_file():
                messagebox.showerror("文件不可用", "请选择有效的本地 PDF、DOCX、PNG、JPG 或 JPEG 文件。")
                return
            self.source.set("其他文件")
            self.file = path
            self.clear()
            self.status.configure(text=f"已选择 {self.file.name}，请点击提取印章")

    def extract(self):
        path = self.current_file() if self.source.get() == "当前申请表 PDF" else self.file
        self.clear()
        if path is None:
            self.status.configure(text="请先选择相应来源的文件")
            return
        if self.source.get() == "当前申请表 PDF" and path.suffix.lower() != ".pdf":
            self.status.configure(text="当前申请表来源需要 PDF；独立图片请使用其他文件来源")
            return
        token = self.generation
        queue = Queue()
        self.status.configure(text="正在读取本地图片……")

        def work():
            try:
                pdf_source = None
                if path.suffix.lower() == ".pdf":
                    started = perf_counter()
                    pdf_source = self.pdf_reader.find(path,
                        lambda text: queue.put(("progress", text)),
                        lambda: token != self.generation)
                    if pdf_source is None:
                        return
                    print(f"[印章按需查找] elapsed={perf_counter() - started:.3f}s cached={self.pdf_reader.last_cache_hit} "
                          f"pending={len(pdf_source.pending_pages)} "
                          + ("ocr_calls=0（本次复用缓存）" if self.pdf_reader.last_cache_hit else
                             " ".join(f"{key}={value}" for key, value in pdf_source.timings.items())))
                    result = ([], pdf_source.notice)
                elif path.suffix.lower() == ".docx":
                    result = extract_docx_images(path)
                elif path.suffix.lower() in {".png", ".jpg", ".jpeg"}:
                    original = candidate_from_bytes(path.read_bytes(), "图片原图，可直接使用或框选裁剪")
                    result = ([original], "已读取原始图片并默认勾选，可直接生成 Word；需要时使用框选裁剪")
                else:
                    raise ValueError()
                queue.put(("done", (result, pdf_source)))
            except PdfPageError as exc:
                queue.put(("done", (([], str(exc)), None)))
            except Exception:
                queue.put(("done", (([], "提取失败，文件无法打开、PDF 无法读取或图片对象不受支持；可重新选择文件"), None)))

        def poll():
            if not self.winfo_exists() or token != self.generation:
                return
            while True:
                try:
                    kind, payload = queue.get_nowait()
                except Empty:
                    self.after(100, poll)
                    return
                if kind == "progress":
                    self.status.configure(text=payload)
                    continue
                (self.candidates, notice), self.pdf_source = payload
                break
            if self.pdf_source is not None:
                self._install_pdf_source(self.pdf_source)
            if path.suffix.lower() in {".png", ".jpg", ".jpeg"} and self.candidates:
                self.selected = [0]
            self.status.configure(text=notice)
            self.render()

        threading.Thread(target=work, daemon=True).start()
        self.after(100, poll)

    def render(self):
        started = perf_counter()
        for widget in self.cards.winfo_children():
            widget.destroy()
        if self.pdf_source is not None:
            manual = ctk.CTkFrame(self.cards)
            manual.pack(fill="x", pady=3)
            ctk.CTkLabel(manual, text=f"PDF 共 {self.pdf_source.page_count} 页，手动页码").pack(side="left", padx=4)
            page_entry = ctk.CTkEntry(manual, width=65)
            page_entry.pack(side="left", padx=4)
            ctk.CTkButton(manual, text="查看 / 框选", width=100,
                          command=lambda: self.open_pdf_page(page_entry.get())).pack(side="left", padx=4)
        for index in range(len(self.candidates)):
            candidate = self.candidates[index]
            card = ctk.CTkFrame(self.cards)
            card.pack(fill="x", pady=3)
            preview = self.thumbnail_cache.get(candidate.data)
            if preview is None:
                with Image.open(BytesIO(candidate.data)) as image:
                    image.thumbnail((90, 60))
                    preview = ctk.CTkImage(image.copy(), size=image.size)
                self.thumbnail_cache[candidate.data] = preview
            thumbnail = ctk.CTkLabel(card, text="", image=preview, width=100, height=65)
            thumbnail.pack(side="left")
            thumbnail.bind("<Button-1>", lambda _event, i=index: self.crop(i))
            order_text = f" · 填入 Word 第 {sorted(self.selected).index(index) + 1} 格" if index in self.selected else ""
            if candidate.selectable:
                check = ctk.CTkCheckBox(card, text=f"候选 {index + 1}{order_text}", command=lambda i=index: self.toggle(i))
                if index in self.selected:
                    check.select()
                check.pack(side="left", padx=4)
            else:
                ctk.CTkLabel(card, text=f"原图 {index + 1}（需框选）").pack(side="left", padx=4)
            ctk.CTkButton(card, text="预览", width=55, command=lambda i=index: self.preview(i)).pack(side="left", padx=3)
            ctk.CTkButton(card, text="框选", width=55, command=lambda i=index: self.crop(i)).pack(side="left", padx=3)
            for label, step in (("上移", -1), ("下移", 1)):
                if 0 <= index + step < len(self.candidates):
                    ctk.CTkButton(card, text=label, width=55,
                                  command=lambda i=index, s=step: self.move(i, s)).pack(side="left", padx=3)
            ctk.CTkLabel(card, text=candidate.context, wraplength=200).pack(side="left", padx=4)
        print(f"[性能统计] 印章候选UI: count={len(self.candidates)} elapsed={perf_counter() - started:.3f}s（UI阶段，不含OCR）")

    def _install_pdf_source(self, source):
        self.pdf_source = source
        existing = {candidate.pdf_page for candidate in self.candidates}
        for page in source.matches:
            if page not in existing and page in source.thumbnails:
                candidate = candidate_from_bytes(source.thumbnails[page], f"PDF 第 {page} 页 · 待框选")
                self.candidates.append(replace(candidate, selectable=False, position=(page, 0, 0), pdf_page=page))

    def open_pdf_page(self, value, index=None):
        if self.pdf_source is None or self.pdf_loading:
            return
        try:
            page = int(value)
            if not 1 <= page <= self.pdf_source.page_count:
                raise ValueError()
        except (ValueError, TypeError):
            self.status.configure(text=f"请输入 1 至 {self.pdf_source.page_count} 之间的有效页码")
            return
        if index is None:
            index = next((i for i, candidate in enumerate(self.candidates) if candidate.pdf_page == page), None)
        if index is not None and self.candidates[index].source_data is not None:
            try:
                if file_signature(self.pdf_source.path) != self.pdf_source.signature:
                    raise OSError()
            except OSError:
                self.status.configure(text="PDF 无法读取或已经变化，请重新提取页面")
                return
            self.crop(index)
            return
        token = self.generation
        source = self.pdf_source
        queue = Queue()
        self.pdf_loading = True
        self.status.configure(text=f"正在高清渲染 PDF 第 {page} 页……")

        def work():
            try:
                queue.put((self.pdf_reader.render(source, page), None))
            except PdfPageError as exc:
                queue.put((None, str(exc)))
            except Exception:
                queue.put((None, "PDF 页面渲染失败，请检查文件是否可读，并尝试其他页码"))

        def poll():
            if not self.winfo_exists() or token != self.generation:
                return
            try:
                candidate, error = queue.get_nowait()
            except Empty:
                self.after(100, poll)
                return
            self.pdf_loading = False
            if error:
                self.status.configure(text=error)
                return
            target = index
            if target is None:
                target = len(self.candidates)
                self.candidates.append(replace(candidate, selectable=False, position=(page, 0, 0),
                                               pdf_page=page, source_data=candidate.data))
            else:
                self.candidates[target] = replace(self.candidates[target], source_data=candidate.data)
            self.status.configure(text=f"PDF 第 {page} 页已渲染，请框选单个印章；整页不会填入 Word")
            self.render()
            self.crop(target)

        threading.Thread(target=work, daemon=True).start()
        self.after(100, poll)

    def toggle(self, index):
        if not self.candidates[index].selectable:
            return
        if index in self.selected:
            self.selected.remove(index)
        elif len(self.selected) < 3:
            self.selected.append(index)
            self.selected.sort()
        else:
            messagebox.showinfo("最多三张", "请先取消一张已选图片，其他候选仍保留。")
        self.status.configure(text=f"已勾选 {len(self.selected)} 张，生成 Word 时按列表顺序填入")
        self.render()

    def move(self, index, step):
        target = index + step
        if not 0 <= target < len(self.candidates):
            return
        self.candidates[index], self.candidates[target] = self.candidates[target], self.candidates[index]
        # 索引随候选整体交换，勾选状态和来源始终跟随图片。
        self.selected = sorted(target if i == index else index if i == target else i for i in self.selected)
        self.status.configure(text="顺序已更新，生成 Word 将按当前列表中的勾选顺序填入")
        self.render()

    def preview(self, index):
        candidate = self.candidates[index]
        if candidate.pdf_page is not None and not candidate.selectable:
            self.open_pdf_page(candidate.pdf_page, index)
            return
        window = ctk.CTkToplevel(self)
        window.title(f"印章候选 {index + 1}，请核对内容与完整性")
        image = Image.open(BytesIO(candidate.data))
        scale = min(1, 850 / image.width, 650 / image.height)
        preview = ctk.CTkImage(image, size=(max(1, int(image.width * scale)), max(1, int(image.height * scale))))
        ctk.CTkLabel(window, text="", image=preview).pack(padx=15, pady=15)
        close = self._modal(window)
        ctk.CTkButton(window, text="关闭", command=close).pack(pady=5)

    def _modal(self, window):
        """窗口映射后再获取焦点和局部 grab，关闭时统一恢复主窗口。"""
        root = self.winfo_toplevel()
        window.transient(root)
        self.windows.append(window)
        pending = []

        def cleanup(event):
            if event.widget != window:
                return
            for callback in pending:
                window.after_cancel(callback)
            try:
                if root.grab_current() == window:
                    window.grab_release()
            except tk.TclError:
                pass
            if window in self.windows:
                self.windows.remove(window)
            if root.winfo_exists():
                root.after_idle(restore_focus)

        def restore_focus():
            if root.winfo_exists() and root.grab_current() is None:
                root.lift()
                root.focus_force()

        def close():
            if window.winfo_exists():
                if window.grab_current() == window:
                    window.grab_release()
                window.destroy()

        def activate():
            if window.winfo_exists() and window.winfo_viewable():
                window.lift()
                window.grab_set()
                window.focus_force()

        def mapped(event):
            if event.widget == window:
                pending.append(window.after_idle(activate))

        window.bind("<Map>", mapped, add="+")
        window.bind("<Destroy>", cleanup, add="+")
        window.bind("<Escape>", lambda _event: close())
        window.protocol("WM_DELETE_WINDOW", close)
        pending.append(window.after(250, activate))
        return close

    def prepare_export(self):
        selected = sorted(self.selected)
        if len(selected) > 3:
            messagebox.showerror("印章数量超限", "最多只能选择三张印章，请取消多余选择。")
            return None
        if any(not self.candidates[i].selectable for i in selected):
            messagebox.showerror("印章尚未框选", "请先框选并保存真实裁剪图片。")
            return None
        return tuple(self.candidates[i].data for i in selected)

    def crop(self, index):
        source = self.candidates[index]
        if source.pdf_page is not None and source.source_data is None:
            self.open_pdf_page(source.pdf_page, index)
            return
        token = self.generation
        original_data = source.source_data or source.data
        original = Image.open(BytesIO(original_data))
        original.load()
        window = ctk.CTkToplevel(self)
        window.title("原图框选：拖动选择单个印章，保留完整外圈和少量白边")
        close = self._modal(window)
        self._center_crop_window(window)
        window.rowconfigure(0, weight=1)
        window.columnconfigure(0, weight=1)
        canvas = tk.Canvas(window, width=900, height=600, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        vertical = tk.Scrollbar(window, orient="vertical", command=canvas.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = tk.Scrollbar(window, orient="horizontal", command=canvas.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        canvas.configure(xscrollcommand=horizontal.set, yscrollcommand=vertical.set)
        scale = min(1, 900 / original.width, 600 / original.height)
        rect = None
        anchor = None
        bounds = None
        photo = None
        space_down = False
        gesture = None
        focus_job = None
        paint_job = None
        zoom_job = None
        zoom_steps = 0.0
        zoom_center = None
        fitted = False
        hand_cursor = "arrow"
        for cursor in ("hand2", "hand1", "fleur"):
            try:
                canvas.configure(cursor=cursor)
                hand_cursor = cursor
                break
            except tk.TclError:
                continue
        canvas.configure(cursor="crosshair")
        status = ctk.CTkLabel(window, text="左键框选 · 空格＋左键平移 · Ctrl＋滚轮缩放")
        status.grid(row=2, column=0)
        result_label = ctk.CTkLabel(window, text="框选后在此显示实际裁剪预览", width=200, height=100)
        result_label.grid(row=0, column=2, padx=8, sticky="n")

        def reset_hand():
            nonlocal space_down, gesture, anchor, zoom_job, zoom_steps
            space_down = False
            gesture = None
            anchor = None
            if zoom_job is not None:
                window.after_cancel(zoom_job)
                zoom_job = None
            zoom_steps = 0.0
            if canvas.winfo_exists():
                canvas.configure(cursor="crosshair")

        def press_space(_event):
            nonlocal space_down, gesture, anchor
            space_down = True
            if gesture == "select":
                gesture = "pan"
                anchor = None
                canvas.scan_mark(canvas.winfo_pointerx() - canvas.winfo_rootx(),
                                 canvas.winfo_pointery() - canvas.winfo_rooty())
            canvas.configure(cursor=hand_cursor)
            return "break"

        def release_space(_event):
            nonlocal space_down
            space_down = False
            canvas.configure(cursor="crosshair")
            return "break"

        def check_focus():
            nonlocal focus_job
            focus_job = None
            if not window.winfo_exists():
                return
            focus = window.focus_displayof()
            if focus is None or not (str(focus) == str(window) or str(focus).startswith(str(window) + ".")):
                reset_hand()

        def lost_focus(_event):
            nonlocal focus_job
            if focus_job is not None:
                window.after_cancel(focus_job)
            focus_job = window.after_idle(check_focus)

        def destroyed(event):
            nonlocal space_down, gesture
            if event.widget == window:
                space_down = False
                gesture = None
                if focus_job is not None:
                    window.after_cancel(focus_job)
                if paint_job is not None:
                    window.after_cancel(paint_job)
                if zoom_job is not None:
                    window.after_cancel(zoom_job)

        def paint():
            nonlocal photo, paint_job
            paint_job = None
            if not canvas.winfo_exists():
                return
            # 仅重绘视口，避免 300 DPI 整页在 500% 时分配巨型图片。
            left = max(0, min(original.width, math.floor(canvas.canvasx(0) / scale)))
            top = max(0, min(original.height, math.floor(canvas.canvasy(0) / scale)))
            right = max(left, min(original.width, math.ceil(canvas.canvasx(canvas.winfo_width()) / scale)))
            bottom = max(top, min(original.height, math.ceil(canvas.canvasy(canvas.winfo_height()) / scale)))
            canvas.delete("viewport")
            if right > left and bottom > top:
                display = original.crop((left, top, right, bottom))
                display = display.resize((max(1, round(display.width * scale)), max(1, round(display.height * scale))))
                photo = ImageTk.PhotoImage(display)
                canvas.create_image(left * scale, top * scale, image=photo, anchor="nw", tags="viewport")
                canvas.tag_lower("viewport")

        def schedule_paint(*_args):
            nonlocal paint_job
            if paint_job is None:
                paint_job = window.after_idle(paint)

        def scrolled(scrollbar, first, last):
            scrollbar.set(first, last)
            schedule_paint()

        canvas.configure(xscrollcommand=lambda first, last: scrolled(horizontal, first, last),
                         yscrollcommand=lambda first, last: scrolled(vertical, first, last))
        def ready(_event=None):
            nonlocal fitted
            if canvas.winfo_width() > 10 and canvas.winfo_height() > 10 and canvas.winfo_viewable():
                if not fitted:
                    fitted = True
                    fit()
                else:
                    schedule_paint()

        canvas.bind("<Configure>", ready)
        canvas.bind("<Map>", ready)

        def draw(new_scale, center=None):
            nonlocal scale, rect
            view_x, view_y = center or (canvas.winfo_width() / 2, canvas.winfo_height() / 2)
            image_x, image_y = canvas.canvasx(view_x) / scale, canvas.canvasy(view_y) / scale
            reset_hand()
            scale = new_scale
            width = max(canvas.winfo_width(), round(original.width * scale))
            height = max(canvas.winfo_height(), round(original.height * scale))
            canvas.configure(scrollregion=(0, 0, width, height))
            canvas.xview_moveto(max(0, image_x * scale - view_x) / width)
            canvas.yview_moveto(max(0, image_y * scale - view_y) / height)
            canvas.delete("selection")
            rect = None
            if bounds:
                rect = canvas.create_rectangle(*(value * scale for value in bounds), outline="red", width=2, tags="selection")
            schedule_paint()

        def fit():
            if canvas.winfo_width() <= 10 or canvas.winfo_height() <= 10:
                return
            draw(min(1, canvas.winfo_width() / original.width, canvas.winfo_height() / original.height))
            canvas.xview_moveto(0)
            canvas.yview_moveto(0)

        def apply_zoom():
            nonlocal zoom_job, zoom_steps
            zoom_job = None
            if gesture is None and zoom_steps:
                minimum = min(0.02, canvas.winfo_width() / original.width, canvas.winfo_height() / original.height)
                factor = 1.12 ** max(-12, min(12, zoom_steps))
                draw(max(minimum, min(5, scale * factor)), zoom_center)
            zoom_steps = 0.0

        def wheel(event):
            nonlocal zoom_job, zoom_steps, zoom_center
            if event.state & 0x4:
                if gesture is None and event.delta:
                    zoom_steps += event.delta / 120
                    zoom_center = (event.x, event.y)
                    if zoom_job is None:
                        zoom_job = window.after(16, apply_zoom)
            elif event.delta:
                canvas.yview_scroll(-int(event.delta / 120) or (-1 if event.delta > 0 else 1), "units")
            return "break"

        def start(event):
            nonlocal anchor, rect, bounds, gesture
            canvas.focus_set()
            if space_down:
                gesture = "pan"
                canvas.scan_mark(event.x, event.y)
                return "break"
            gesture = "select"
            anchor = (canvas.canvasx(event.x), canvas.canvasy(event.y))
            if rect is not None:
                canvas.delete(rect)
            rect = canvas.create_rectangle(*anchor, *anchor, outline="red", width=2, tags="selection")
            bounds = None
            result_label.configure(image=None, text="拖动选择印章")

        def drag(event):
            nonlocal bounds
            if gesture == "pan":
                if space_down:
                    canvas.scan_dragto(event.x, event.y, gain=1)
                return "break"
            if gesture != "select" or space_down:
                return "break"
            if anchor is None or rect is None:
                return
            x, y = canvas.canvasx(event.x), canvas.canvasy(event.y)
            canvas.coords(rect, *anchor, x, y)
            bounds = (max(0, math.floor(min(anchor[0], x) / scale)),
                      max(0, math.floor(min(anchor[1], y) / scale)),
                      min(original.width, math.ceil(max(anchor[0], x) / scale)),
                      min(original.height, math.ceil(max(anchor[1], y) / scale)))
            status.configure(text=f"原始像素裁剪坐标 {bounds}，保存后生成真实裁剪图片")

        def show_result(event):
            nonlocal gesture, anchor
            if gesture != "select" or space_down:
                gesture = None
                anchor = None
                return "break"
            drag(event)
            gesture = None
            anchor = None
            if bounds and bounds[2] > bounds[0] and bounds[3] > bounds[1]:
                cropped = original.crop(bounds)
                ratio = min(200 / cropped.width, 200 / cropped.height)
                image = ctk.CTkImage(cropped, size=(max(1, round(cropped.width * ratio)), max(1, round(cropped.height * ratio))))
                result_label.configure(image=image, text="")

        def save(as_new=False):
            if token != self.generation or not bounds or bounds[2] <= bounds[0] or bounds[3] <= bounds[1]:
                return
            output = BytesIO()
            try:
                cropped = original.crop(bounds)
                if cropped.mode not in {"1", "L", "LA", "P", "RGB", "RGBA", "I", "I;16"}:
                    messagebox.showinfo("图片颜色模式不支持", "当前图片颜色模式无法无损保存为 PNG，请使用常规 RGB 图片后框选。", parent=window)
                    return
                cropped.save(output, format="PNG")
                candidate = candidate_from_bytes(output.getvalue(), "人工框选的实际裁剪图片")
            except Exception:
                messagebox.showerror("裁剪失败", "未能保存实际裁剪图片，请重新框选或使用常规 PNG/JPG 原图。", parent=window)
                return
            candidate = replace(candidate, position=(source.position[0], bounds[1], bounds[0]), source_data=original_data)
            candidate = replace(candidate, pdf_page=source.pdf_page)
            if source.pdf_page is not None:
                candidate = replace(candidate, context=f"PDF 第 {source.pdf_page} 页人工裁剪")
            needs_selection = as_new or index not in self.selected
            if needs_selection and len(self.selected) >= 3:
                messagebox.showinfo("最多三张", "当前已选择三张，请取消一张后再保存；本次尚未提交。", parent=window)
                return
            self._commit_crop(index, candidate, as_new)
            close()

        buttons = ctk.CTkFrame(window)
        buttons.grid(row=3, column=0, pady=5)
        ctk.CTkButton(buttons, text="原始像素 1:1", width=110, command=lambda: draw(1)).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="适应窗口", width=100, command=fit).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="保存裁剪", width=100, command=save).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="另存为新印章", width=110, command=lambda: save(True)).pack(side="left", padx=4)
        ctk.CTkButton(buttons, text="取消", width=80, command=close).pack(side="left", padx=4)
        canvas.bind("<ButtonPress-1>", start)
        canvas.bind("<B1-Motion>", drag)
        canvas.bind("<ButtonRelease-1>", show_result)
        canvas.bind("<MouseWheel>", wheel)
        window.bind("<KeyPress-space>", press_space, add="+")
        window.bind("<KeyRelease-space>", release_space, add="+")
        window.bind("<FocusOut>", lost_focus, add="+")
        window.bind("<Destroy>", destroyed, add="+")
        window.after_idle(ready)

    def _commit_crop(self, index, candidate, as_new):
        """保存默认替换当前候选，只有显式另存才新增。"""
        if as_new:
            index += 1
            self.candidates.insert(index, candidate)
            self.selected = [i + 1 if i >= index else i for i in self.selected]
        else:
            self.candidates[index] = candidate
        if index not in self.selected:
            self.selected.append(index)
            self.selected.sort()
        self.status.configure(text=f"裁剪已保存并勾选，共 {len(self.selected)} 张，生成 Word 时直接填入")
        self.render()

    def _center_crop_window(self, window):
        """按主窗口所在显示器工作区居中，给边框和任务栏留出空间。"""
        root = self.winfo_toplevel()
        left, top = root.winfo_vrootx(), root.winfo_vrooty()
        right, bottom = left + root.winfo_vrootwidth(), top + root.winfo_vrootheight()
        try:
            import ctypes
            from ctypes import wintypes

            class MonitorInfo(ctypes.Structure):
                _fields_ = [("cbSize", wintypes.DWORD), ("monitor", wintypes.RECT),
                            ("work", wintypes.RECT), ("flags", wintypes.DWORD)]

            user32 = ctypes.windll.user32
            user32.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
            user32.MonitorFromWindow.restype = wintypes.HANDLE
            user32.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MonitorInfo)]
            info = MonitorInfo()
            info.cbSize = ctypes.sizeof(info)
            monitor = user32.MonitorFromWindow(root.winfo_id(), 2)
            if user32.GetMonitorInfoW(monitor, ctypes.byref(info)):
                left, top, right, bottom = info.work.left, info.work.top, info.work.right, info.work.bottom
        except (AttributeError, OSError):
            pass
        width, height = min(1000, right - left - 32), min(650, bottom - top - 64)
        x = max(left + 8, min(right - width - 16, root.winfo_rootx() + (root.winfo_width() - width) // 2))
        y = max(top + 8, min(bottom - height - 48, root.winfo_rooty() + (root.winfo_height() - height) // 2))
        scaling = window._get_window_scaling()
        window.geometry(f"{round(width / scaling)}x{round(height / scaling)}+{x}+{y}")
