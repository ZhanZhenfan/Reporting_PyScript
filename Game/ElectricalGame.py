import os
import random
import sys
import time
import tkinter as tk
from tkinter import messagebox

try:
    from PIL import Image, ImageTk  # type: ignore
    PIL_OK = True
except Exception as e:
    PIL_OK = False
    PIL_ERR = e

# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

# Singapore electricity tariff (SGD per kWh).  Updated for Q4 2025.
SG_TARIFF_SGD_PER_KWH = 0.2755

# Equipment definitions. Each entry: appliance name (EN) -> (wattage, hours_per_day).
EQUIPMENT: dict[str, tuple[int, float]] = {
    "Central Air Conditioner": (3500, 12.0),
    "Microwave": (800, 1.0),
    "Water Dispenser (Heating)": (650, 4.0),
    "Refrigerator": (500, 8.0),
    "Printer (Printing)": (400, 2.0),
    "Projector": (250, 12.0),
    "LED Tube": (18, 12.0),
    "Desktop PC": (150, 8.0),
    "Laptop": (45, 8.0),
}

SCENARIOS_EN: dict[str, str] = {
    "Central Air Conditioner": "Left on after work (12 h/day)",
    "Microwave": "Multiple heating sessions (1 h/day)",
    "Water Dispenser (Heating)": "Overheating due to excess use (4 h/day)",
    "Refrigerator": "Compressor runs to keep cold (8 h/day)",
    "Printer (Printing)": "Long printing tasks (2 h/day)",
    "Projector": "Not turned off after meeting (12 h/day)",
    "LED Tube": "Lights left on when room is empty (12 h/day)",
    "Desktop PC": "Idle or unused for hours (8 h/day)",
    "Laptop": "Charging/standby for hours (8 h/day)",
}

SCENARIOS_ZH: dict[str, str] = {
    "Central Air Conditioner": "下班后仍未关闭（约 12 小时/天）",
    "Microwave": "一天多次加热食物（约 1 小时/天）",
    "Water Dispenser (Heating)": "频繁加热、长时间保温（约 4 小时/天）",
    "Refrigerator": "压缩机周期运行保冷（约 8 小时/天）",
    "Printer (Printing)": "长时间连续打印（约 2 小时/天）",
    "Projector": "会议结束后忘记关机（约 12 小时/天）",
    "LED Tube": "人离开房间灯仍常亮（约 12 小时/天）",
    "Desktop PC": "空闲时长时间不关机（约 8 小时/天）",
    "Laptop": "充满电仍插着电源待机（约 8 小时/天）",
}

NAME_ZH: dict[str, str] = {
    "Central Air Conditioner": "中央空调",
    "Microwave": "微波炉",
    "Water Dispenser (Heating)": "冷热饮水机（加热）",
    "Refrigerator": "冰箱",
    "Printer (Printing)": "打印机（打印中）",
    "Projector": "投影仪",
    "LED Tube": "LED 日光灯管",
    "Desktop PC": "台式电脑",
    "Laptop": "笔记本电脑",
}

SUGGESTIONS_EN: dict[str, str] = {
    "Central Air Conditioner": "Turn off or adjust thermostat when you leave.",
    "Microwave": "Use only when necessary and unplug when not in use.",
    "Water Dispenser (Heating)": "Heat only the water you need and turn off heating.",
    "Refrigerator": "Minimise door opening and maintain proper temperature.",
    "Printer (Printing)": "Print only when necessary and turn off afterwards.",
    "Projector": "Turn off immediately after the meeting or set an auto-off timer.",
    "LED Tube": "Turn off lights when you leave or install motion sensors.",
    "Desktop PC": "Shut down or use sleep mode when idle.",
    "Laptop": "Unplug charger when full and power down when not in use.",
}

SUGGESTIONS_ZH: dict[str, str] = {
    "Central Air Conditioner": "离开前关闭空调，或适当调高设定温度。",
    "Microwave": "只在需要时使用，用完及时关闭并拔掉插头。",
    "Water Dispenser (Heating)": "只加热所需水量，长时间不用时关闭加热功能。",
    "Refrigerator": "减少开门次数，合理设置温度并避免塞得太满。",
    "Printer (Printing)": "只打印必要文件，打印结束后关闭打印机电源。",
    "Projector": "会议结束立即关机，或设置自动关机定时。",
    "LED Tube": "人走关灯，或安装感应开关自动控制照明。",
    "Desktop PC": "空闲时启用睡眠模式或直接关机。",
    "Laptop": "充满电后拔掉充电器，长时间不用时关机。",
}

IMAGE_FILENAMES = {
    "Central Air Conditioner": "central_air_conditioner.jpg",
    "Microwave": "microwave.jpg",
    "Water Dispenser (Heating)": "water_dispenser_heating.jpg",
    "Refrigerator": "refrigerator.jpg",
    "Printer (Printing)": "printer_printing.jpg",
    "Projector": "projector.jpg",
    "LED Tube": "led_tubes.jpg",
    "Desktop PC": "desktop_pc.jpg",
    "Laptop": "laptop.jpg",
}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGE_PATHS = {name: os.path.join(BASE_DIR, fname) for name, fname in IMAGE_FILENAMES.items()}


class DragDropGame:
    """
    Drag-and-drop ranking game with language toggle (English / 中文).
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.geometry("1450x780")

        self.lang: str = "en"
        self.num_cards: int = 4

        self.sample: list[tuple[str, tuple[int, float]]] = []
        self.correct_order: list[str] = []
        self.slot_centers: list[tuple[float, float]] = []
        self.card_for_slot: dict[int, str] = {}
        self.slot_for_card: dict[str, int] = {}
        self.cards: dict[str, dict[str, list[int] | int]] = {}
        self.dragging_tag: str | None = None
        self.drag_prev: tuple[int, int] = (0, 0)
        self.last_motion_ts = 0.0
        self.already_revealed: bool = False
        self.image_cache: dict[str, ImageTk.PhotoImage] = {}

        # Layout (unchanged)
        self.front_img_max_w = 320
        self.front_img_max_h = 320
        self.left_margin = 30
        self.top_margin = 210
        self.slot_w = self.front_img_max_w + 90
        self.slot_h = self.front_img_max_h + 260
        self.slot_gap = 22
        self.card_pad = 12
        self.anim_steps = 8
        self.anim_delay = 12

        # ----- Fonts (SUGGESTION enlarged & bolder) -----
        self.font_front_name = ("Segoe UI", 15, "bold")
        self.font_front_scenario = ("Segoe UI", 13, "bold")
        self.font_back_title = ("Segoe UI", 15, "bold")
        self.font_back_cost = ("Segoe UI", 22, "bold")
        self.font_back_hint = ("Segoe UI", 11)
        # ↑↑ 放大到 20，并保持加粗
        self.font_back_suggestion = ("Segoe UI", 20, "bold")

        # Header & buttons
        self.header_lbl = tk.Label(root, text="", font=("Segoe UI", 16, "bold"))
        self.header_lbl.pack(pady=(10, 6))

        tb = tk.Frame(root)
        tb.pack()

        self.btn_draw = tk.Button(tb, text="", command=self.start_game, width=14)
        self.btn_draw.pack(side="left", padx=6)

        self.btn_check = tk.Button(tb, text="", command=self.on_check_and_reveal, width=14)
        self.btn_check.pack(side="left", padx=6)

        self.btn_reset = tk.Button(tb, text="", command=self.reset_board, width=14)
        self.btn_reset.pack(side="left", padx=6)

        self.btn_toggle_lang = tk.Button(tb, text="", command=self.toggle_language, width=10)
        self.btn_toggle_lang.pack(side="left", padx=12)

        self.canvas = tk.Canvas(root, bg="#f7f8fb", highlightthickness=0)
        self.canvas.pack(expand=True, fill="both", padx=16, pady=10)

        self.hint_id: int | None = None
        self.tariff_id: int | None = None

        self.update_static_labels()
        self.reset_board()

        self.canvas.bind("<Configure>", self.on_resize)
        self.canvas.tag_bind("card", "<ButtonPress-1>", self.on_press)
        self.canvas.tag_bind("card", "<B1-Motion>", self.on_drag)
        self.canvas.tag_bind("card", "<ButtonRelease-1>", self.on_release)

        self._log_missing_images()
        if not PIL_OK:
            messagebox.showwarning(
                self._tr("Pillow missing", "缺少 Pillow 库"),
                self._tr(
                    "Pillow is required for JPG images.\nInstall with:\n\npip install pillow\n\n",
                    "加载 JPG 图片需要 Pillow 库。\n请通过以下命令安装：\n\npip install pillow\n\n"
                ) + str(PIL_ERR)
            )

    # ---------------- Language helpers ----------------
    def _tr(self, en: str, zh: str) -> str:
        return en if self.lang == "en" else zh

    def update_static_labels(self) -> None:
        if self.lang == "en":
            self.root.title("Smart Saver Ranking Game (Highest → Lowest Annual Cost)")
        else:
            self.root.title("Smart Saver 省电排序游戏（年度电费从高到低）")

        self.header_lbl.config(text=self._tr(
            "Drag to rank (Left → Right) Highest → → → Lowest annual cost",
            "拖动卡牌从左到右排序：年度电费从高到低"
        ))

        self.btn_draw.config(text=self._tr("Draw 4 Cards", "抽取 4 张卡牌"))
        self.btn_check.config(text=self._tr("Check Answer", "检查答案"))
        self.btn_reset.config(text=self._tr("Reset", "重新开始"))
        self.btn_toggle_lang.config(text=self._tr("中文", "English"))

    def toggle_language(self) -> None:
        self.lang = "zh" if self.lang == "en" else "en"
        self.update_static_labels()
        if not self.sample:
            self.reset_board()
        else:
            order_names: list[str | None] = []
            for i in range(self.num_cards):
                tag = self.card_for_slot.get(i)
                order_names.append(tag.replace("card_name_", "") if tag else None)
            self.layout_slots_and_cards(
                keep_order=order_names, keep_face=self.already_revealed
            )

    # ---------------- Utilities ----------------
    def _log_missing_images(self) -> None:
        missing = [(n, os.path.basename(p)) for n, p in IMAGE_PATHS.items() if not os.path.exists(p)]
        if missing:
            print(">> Missing image files (place these .jpg in the same folder):")
            for n, fname in missing:
                print(f"   - {n}  ->  {fname}")
        else:
            print(">> All standard JPG filenames found.")

    def _hint_text(self) -> str:
        return self._tr(
            "Place cards Left → Right  (Highest → Lowest annual cost)",
            "从左到右放置卡牌：年度电费从高到低"
        )

    def _tariff_text(self) -> str:
        if self.lang == "en":
            return (
                f"Annual cost = watts × hours/day × 365 × tariff  •  "
                f"Tariff: S${SG_TARIFF_SGD_PER_KWH:.4f}/kWh (Q4 2025)"
            )
        else:
            return (
                f"年电费 = 功率(瓦) ÷ 1000 × 每日小时数 × 365 天 × 电价  •  "
                f"电价：约 S${SG_TARIFF_SGD_PER_KWH:.4f}/千瓦时（2025 年第 4 季度）"
            )

    def reset_board(self) -> None:
        self.canvas.delete("all")
        w = self.canvas.winfo_width() or 1400
        self.hint_id = self.canvas.create_text(
            w / 2, 90, text=self._hint_text(),
            font=("Segoe UI", 12), fill="#666"
        )
        self.tariff_id = self.canvas.create_text(
            w / 2, 120, text=self._tariff_text(),
            font=("Segoe UI", 10, "bold"), fill="#666"
        )
        self.sample.clear()
        self.correct_order.clear()
        self.slot_centers.clear()
        self.card_for_slot.clear()
        self.slot_for_card.clear()
        self.cards.clear()
        self.already_revealed = False

    def start_game(self) -> None:
        self.reset_board()
        self.sample = random.sample(list(EQUIPMENT.items()), self.num_cards)
        cost_list: list[tuple[str, float]] = []
        for name, (watts, hours) in self.sample:
            kwh_year = (watts / 1000.0) * (hours * 365)
            cost_year = kwh_year * SG_TARIFF_SGD_PER_KWH
            cost_list.append((name, cost_year))
        self.correct_order = [n for n, _ in sorted(cost_list, key=lambda x: x[1], reverse=True)]
        self.layout_slots_and_cards()

    # ---------------- Layout ----------------
    def on_resize(self, event: tk.Event) -> None:
        if not self.sample:
            self.center_hint_only()
            return
        order_names: list[str | None] = []
        for i in range(self.num_cards):
            tag = self.card_for_slot.get(i)
            order_names.append(tag.replace("card_name_", "") if tag else None)
        self.layout_slots_and_cards(
            keep_order=order_names,
            keep_face=self.already_revealed
        )

    def center_hint_only(self) -> None:
        w = self.canvas.winfo_width() or 1400
        if self.hint_id is not None:
            self.canvas.coords(self.hint_id, w / 2, 90)
        if self.tariff_id is not None:
            self.canvas.coords(self.tariff_id, w / 2, 120)

    def layout_slots_and_cards(
        self,
        keep_order: list[str | None] | None = None,
        keep_face: bool = False
    ) -> None:
        self.canvas.delete("all")
        w = self.canvas.winfo_width() or 1400

        self.hint_id = self.canvas.create_text(
            w / 2, 90, text=self._tr(
                "Left → Right = Highest → Lowest annual cost",
                "从左到右：年度电费从高到低"
            ),
            font=("Segoe UI", 12), fill="#666"
        )
        self.tariff_id = self.canvas.create_text(
            w / 2, 120, text=self._tariff_text(),
            font=("Segoe UI", 10, "bold"), fill="#666"
        )

        total_min = self.num_cards * self.slot_w + (self.num_cards - 1) * self.slot_gap
        usable_w = max(total_min, min(w - 2 * self.left_margin, 2200))
        gap = self.slot_gap
        if usable_w > total_min and self.num_cards > 1:
            extra = usable_w - total_min
            gap = self.slot_gap + extra / (self.num_cards - 1)

        y1 = self.top_margin
        y2 = y1 + self.slot_h
        self.slot_centers.clear()
        left = (w - (self.num_cards * self.slot_w + (self.num_cards - 1) * gap)) / 2
        for i in range(self.num_cards):
            x1 = left + i * (self.slot_w + gap)
            x2 = x1 + self.slot_w
            self.canvas.create_rectangle(
                x1, y1, x2, y2,
                outline="#b9bfd3", dash=(4, 2), width=2,
                fill="", tags=("slot", f"slot_{i}")
            )
            cx = (x1 + x2) / 2
            cy = (y1 + y2) / 2
            self.slot_centers.append((cx, cy))

        if keep_order and any(keep_order):
            picked: list[tuple[str, tuple[int, float]]] = [(n, EQUIPMENT[n]) for n in keep_order if n]
            if len(picked) < self.num_cards:
                remaining = [n for n in dict(self.sample).keys() if n not in keep_order]
                picked += [(n, EQUIPMENT[n]) for n in remaining]
        else:
            picked = list(self.sample)

        self.cards.clear()
        self.card_for_slot.clear()
        self.slot_for_card.clear()

        for i, (name, (watts, hours)) in enumerate(picked[:self.num_cards]):
            cx, cy = self.slot_centers[i]
            half_w = (self.slot_w / 2) - self.card_pad
            half_h = (self.slot_h / 2) - self.card_pad

            rect = self.canvas.create_rectangle(
                cx - half_w, cy - half_h, cx + half_w, cy + half_h,
                fill="#ffffff", outline="#5b8def", width=2,
                tags=("card", f"card_name_{name}")
            )

            img_ids = self.create_front_image(name, cx, cy - 40)

            txt_ids: list[int] = []
            text_y = cy + (self.front_img_max_h / 2) - 8

            def add_text_line(text: str, font, fill: str = "#000") -> None:
                nonlocal text_y, txt_ids
                if not text:
                    return
                tid = self.canvas.create_text(
                    cx, text_y,
                    text=text,
                    font=font,
                    fill=fill,
                    width=self.slot_w - 2 * self.card_pad - 20,
                    tags=("card", f"card_name_{name}")
                )
                bbox = self.canvas.bbox(tid)
                if bbox:
                    _, _, _, bottom = bbox
                    text_y = bottom + 4
                else:
                    text_y += 20
                txt_ids.append(tid)

            if self.lang == "en":
                title = name
                scenario = SCENARIOS_EN.get(name, "")
            else:
                title = NAME_ZH.get(name, name)
                scenario = SCENARIOS_ZH.get(name, "")

            add_text_line(title, self.font_front_name, "#111")
            add_text_line(scenario, self.font_front_scenario, "#555")

            self.cards[f"card_name_{name}"] = {
                "rect": rect,
                "img_ids": img_ids,
                "txt_ids": txt_ids,
            }
            self.card_for_slot[i] = f"card_name_{name}"
            self.slot_for_card[f"card_name_{name}"] = i

            if keep_face:
                self.show_back(f"card_name_{name}", name)

        self.canvas.tag_bind("card", "<ButtonPress-1>", self.on_press)
        self.canvas.tag_bind("card", "<B1-Motion>", self.on_drag)
        self.canvas.tag_bind("card", "<ButtonRelease-1>", self.on_release)

    # ---------------- Image loading ----------------
    def load_photo(self, path: str, max_w: int, max_h: int) -> ImageTk.PhotoImage | None:
        if not PIL_OK or not path or not os.path.isfile(path):
            return None
        try:
            im = Image.open(path)  # type: ignore[attr-defined]
            im = im.convert("RGBA")
            im.thumbnail((max_w, max_h), Image.LANCZOS)
            return ImageTk.PhotoImage(im)  # type: ignore[attr-defined]
        except Exception as e:
            print(f">> Failed to load image: {path}\n   {e}")
            return None

    def create_front_image(self, name: str, cx: float, cy: float) -> list[int]:
        ids: list[int] = []
        path = IMAGE_PATHS.get(name)
        img_obj = self.load_photo(path, self.front_img_max_w, self.front_img_max_h)
        if img_obj is not None:
            self.image_cache[name] = img_obj
            ids.append(self.canvas.create_image(
                cx, cy, image=img_obj, tags=("card", f"card_name_{name}")
            ))
        else:
            half_w = self.front_img_max_w // 2
            half_h = self.front_img_max_h // 2
            ids.append(self.canvas.create_rectangle(
                cx - half_w, cy - half_h, cx + half_w, cy + half_h,
                fill="#eef2ff", outline="#9fb0e9", width=1,
                tags=("card", f"card_name_{name}")
            ))
            ids.append(self.canvas.create_text(
                cx, cy, text="🖼", font=("Segoe UI Emoji", 22),
                tags=("card", f"card_name_{name}")
            ))
        return ids

    # ---------------- Sound feedback ----------------
    def play_correct_sound(self) -> None:
        try:
            if sys.platform.startswith("win"):
                import winsound  # type: ignore
                winsound.Beep(1000, 200)
            else:
                self.root.bell()
        except Exception:
            pass

    def play_wrong_sound(self) -> None:
        try:
            if sys.platform.startswith("win"):
                import winsound  # type: ignore
                winsound.Beep(400, 500)
            else:
                self.root.bell()
        except Exception:
            pass

    def flash_cards(self, correct: bool) -> None:
        color = "#28a745" if correct else "#dc3545"
        for tag in self.card_for_slot.values():
            rect_id = self.cards[tag]["rect"]  # type: ignore[index]
            self.canvas.itemconfig(rect_id, outline=color, width=3)

        def revert() -> None:
            for tag in self.card_for_slot.values():
                rect_id = self.cards[tag]["rect"]  # type: ignore[index]
                self.canvas.itemconfig(rect_id, outline="#5b8def", width=2)

        self.canvas.after(300, revert)

    # ---------------- Drag handling ----------------
    def on_press(self, event: tk.Event) -> None:
        item = self.canvas.find_closest(event.x, event.y)[0]
        tags = self.canvas.gettags(item)
        tag = None
        for t in tags:
            if t.startswith("card_name_"):
                tag = t
                break
        if not tag:
            return
        self.dragging_tag = tag
        self.drag_prev = (event.x, event.y)
        card = self.cards[tag]
        self.canvas.tag_raise(card["rect"])
        for iid in card["img_ids"]:  # type: ignore[index]
            self.canvas.tag_raise(iid)
        for tid in card["txt_ids"]:  # type: ignore[index]
            self.canvas.tag_raise(tid)

    def on_drag(self, event: tk.Event) -> None:
        if not self.dragging_tag:
            return
        now = time.time()
        if now - self.last_motion_ts < 0.01:
            return
        dx = event.x - self.drag_prev[0]
        dy = event.y - self.drag_prev[1]
        if dx or dy:
            card = self.cards[self.dragging_tag]
            self.canvas.move(card["rect"], dx, dy)
            for iid in card["img_ids"]:  # type: ignore[index]
                self.canvas.move(iid, dx, dy)
            for tid in card["txt_ids"]:  # type: ignore[index]
                self.canvas.move(tid, dx, dy)
            self.drag_prev = (event.x, event.y)
            self.last_motion_ts = now

    def on_release(self, event: tk.Event) -> None:
        if not self.dragging_tag:
            return
        cx, cy = self.card_center(self.dragging_tag)
        slot_idx = self.nearest_slot(cx, cy)
        old_slot = self.slot_for_card.get(self.dragging_tag)
        occupant = self.card_for_slot.get(slot_idx)

        if occupant and occupant != self.dragging_tag:
            if old_slot is not None:
                self.card_for_slot[old_slot] = occupant
                self.slot_for_card[occupant] = old_slot
                self.animate_to(occupant, self.slot_centers[old_slot])
            else:
                for k in range(self.num_cards):
                    if k not in self.card_for_slot:
                        self.card_for_slot[k] = occupant
                        self.slot_for_card[occupant] = k
                        self.animate_to(occupant, self.slot_centers[k])
                        break

        self.card_for_slot[slot_idx] = self.dragging_tag
        self.slot_for_card[self.dragging_tag] = slot_idx
        self.animate_to(self.dragging_tag, self.slot_centers[slot_idx])
        self.dragging_tag = None

    # ---------------- Helpers ----------------
    def card_center(self, tag: str) -> tuple[float, float]:
        rect_id = self.cards[tag]["rect"]  # type: ignore[index]
        x1, y1, x2, y2 = self.canvas.coords(rect_id)
        return ((x1 + x2) / 2, (y1 + y2) / 2)

    def animate_to(self, tag: str, target_xy: tuple[float, float]) -> None:
        cx, cy = self.card_center(tag)
        tx, ty = target_xy
        dx = (tx - cx) / self.anim_steps
        dy = (ty - cy) / self.anim_steps

        def step(k: int = 0) -> None:
            if k >= self.anim_steps:
                cx2, cy2 = self.card_center(tag)
                self.move_card_exact(tag, tx - cx2, ty - cy2)
                return
            self.move_card_exact(tag, dx, dy)
            self.canvas.after(self.anim_delay, step, k + 1)

        step()

    def move_card_exact(self, tag: str, dx: float, dy: float) -> None:
        card = self.cards[tag]
        self.canvas.move(card["rect"], dx, dy)
        for iid in card["img_ids"]:  # type: ignore[index]
            self.canvas.move(iid, dx, dy)
        for tid in card["txt_ids"]:  # type: ignore[index]
            self.canvas.move(tid, dx, dy)

    def nearest_slot(self, x: float, y: float) -> int:
        best = 0
        best_d = float("inf")
        for i, (sx, sy) in enumerate(self.slot_centers):
            d = (sx - x) ** 2 + (sy - y) ** 2
            if d < best_d:
                best_d = d
                best = i
        return best

    # ---------------- Check & Reveal ----------------
    def on_check_and_reveal(self) -> None:
        if not self.sample:
            return
        names: list[str] = []
        for i in range(self.num_cards):
            tag = self.card_for_slot.get(i)
            if not tag:
                messagebox.showinfo(
                    self._tr("Result", "结果"),
                    self._tr(
                        "Please place all cards into slots first.",
                        "请先把所有卡牌都放入下面的格子中。"
                    )
                )
                return
            names.append(tag.replace("card_name_", ""))

        is_correct = names == self.correct_order
        if is_correct:
            self.play_correct_sound()
        else:
            self.play_wrong_sound()
        self.flash_cards(is_correct)

        if not self.already_revealed:
            for i in range(self.num_cards):
                tag = self.card_for_slot[i]
                name = tag.replace("card_name_", "")
                self.show_back(tag, name)
            self.already_revealed = True
        else:
            messagebox.showinfo(
                self._tr("Shown", "已显示"),
                self._tr(
                    "Answers are already displayed. Use Reset to start a new round.",
                    "答案已经显示。如需重新游戏，请点击“重新开始”。"
                )
            )

    def show_back(self, tag: str, name: str) -> None:
        watts, hours = EQUIPMENT[name]
        kwh_year = (watts / 1000.0) * (hours * 365)
        cost_year = kwh_year * SG_TARIFF_SGD_PER_KWH

        items = self.cards[tag]
        rect_id, img_ids, txt_ids = items["rect"], items["img_ids"], items["txt_ids"]

        # 清理 front 面图文
        for iid in img_ids:
            try:
                self.canvas.delete(iid)
            except Exception:
                pass
        items["img_ids"] = []
        for tid in txt_ids:
            try:
                self.canvas.delete(tid)
            except Exception:
                pass
        items["txt_ids"] = []

        self.canvas.itemconfig(rect_id, fill="#f2f5ff", outline="#3a6be0", width=2)

        x1, y1, x2, y2 = self.canvas.coords(rect_id)
        cx = (x1 + x2) / 2
        cy = (y1 + y2) / 2

        # === 上半部：标题 / 年费用 —— 整体上移，避免和建议重叠 ===
        title_text = name if self.lang == "en" else NAME_ZH.get(name, name)
        title_id = self.canvas.create_text(
            cx, cy - 115,  # ← 上移
            text=title_text,
            font=self.font_back_title,
            fill="#0b3d91",
            tags=("card", tag)
        )

        cost_str = (f"≈ S${cost_year:,.2f} / year"
                    if self.lang == "en"
                    else f"≈ 每年约 S${cost_year:,.2f}")
        cost_id = self.canvas.create_text(
            cx, cy - 75,  # ← 上移
            text=cost_str,
            font=self.font_back_cost,
            fill="#111",
            tags=("card", tag)
        )

        # === 中部偏下：建议 —— 放得更低一点，并水平居中换行 ===
        base_suggestion = (SUGGESTIONS_EN if self.lang == "en" else SUGGESTIONS_ZH).get(name, "")
        suggestion_text = (f"Suggested action: {base_suggestion}"
                           if self.lang == "en"
                           else f"建议行动：{base_suggestion}")

        wrap_width = self.slot_w - 2 * self.card_pad - 30
        suggestion_id = self.canvas.create_text(
            cx, cy + 30,  # ← 往下挪（原来是 +5）
            text=suggestion_text,
            font=self.font_back_suggestion,
            fill="#333",
            width=wrap_width,
            justify="center",
            tags=("card", tag)
        )

        # === 最下：小时提示 —— 再往下放，保证留白 ===
        hint_str = (f"({hours:.0f} h/day × 365 days)"
                    if self.lang == "en"
                    else f"（每天约 {hours:.0f} 小时 × 365 天）")
        hint_id = self.canvas.create_text(
            cx, cy - 45,  # ← 更靠下
            text=hint_str,
            font=self.font_back_hint,
            fill="#777",
            justify="center",
            tags=("card", tag)
        )

        items["txt_ids"] = [title_id, cost_id, suggestion_id, hint_id]


if __name__ == "__main__":
    root_app = tk.Tk()
    game = DragDropGame(root_app)
    root_app.mainloop()
