"""
Modified Water-Saving Q&A (A/B) quiz application.

This version of the quiz adds illustrative photos for each question and
restructures the layout into two distinct regions.  The top region,
colored a light blue reminiscent of clean water, holds the question text
and the choice buttons.  The lower region is used to either display
the relevant photograph before an answer has been submitted or to
show the recommended answer along with an explanation once the
participant has made a choice.  Images are cropped to a consistent
size and placed alongside the script in the same directory.
"""

import os
import random
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

APP_TITLE_EN = "Water-Saving Q&A (A/B)"
APP_TITLE_ZH = "节水问答（A/B）"

FACILITIES_HOTLINE_EN = "8428 9066 or 8428 9039"
FACILITIES_HOTLINE_ZH = "8428 9066 或 8428 9039"

# ===================== 题库 =====================
QUESTION_BANK = [
    {
        "q": {
            "en": "When washing hands with soap, what should you do to save water?",
            "zh": "用肥皂洗手时，为了节水应该怎么做？",
        },
        "A": {
            "en": "Wet hands, then turn OFF the tap while rubbing; turn it ON only to rinse.",
            "zh": "先把手打湿，搓手时关闭水龙头，只在冲洗时再打开。",
        },
        "B": {
            "en": "Keep the tap running the whole time for convenience.",
            "zh": "为了方便，全程开着水龙头。",
        },
        "correct": "A",
        "explain": {
            "en": "Running water is only needed to wet and to rinse. Turning it off during rubbing saves a lot of water.",
            "zh": "流水只在打湿和冲洗时需要；搓手期间关水可显著节约用水。",
        },
    },
    {
        "q": {
            "en": "You see a leaking tap (continuous drip). What should you do?",
            "zh": "看到水龙头漏水或持续滴水，应该怎么处理？",
        },
        "A": {
            "en": "Ignore it for now; someone will fix it eventually.",
            "zh": "先忽略，总会有人来修。",
        },
        "B": {
            "en": f"Report it to Facilities ({FACILITIES_HOTLINE_EN}) immediately.",
            "zh": f"立即报修：联系设施团队（{FACILITIES_HOTLINE_ZH}）。",
        },
        "correct": "B",
        "explain": {
            "en": "Small leaks waste large volumes over time; prompt repair prevents waste and damage.",
            "zh": "小漏长期会浪费大量用水；尽快报修可以避免浪费与设备损坏。",
        },
    },
    {
        "q": {
            "en": "Water-saving habit when brushing your teeth at work?",
            "zh": "在办公室刷牙时的节水习惯是？",
        },
        "A": {
            "en": "Turn off the tap while brushing (or use a cup); run water only to rinse.",
            "zh": "刷牙时关水或用杯子，只有在漱口时开水。",
        },
        "B": {
            "en": "Keep tap at medium flow throughout brushing.",
            "zh": "刷牙全程保持中等水流。",
        },
        "correct": "A",
        "explain": {
            "en": "Keeping the tap running adds up; short, targeted rinsing is enough.",
            "zh": "一直开水会累计大量浪费；短暂、有针对性的冲洗就足够。",
        },
    },
    {
        "q": {
            "en": "What’s a water-saving way to wash mugs/dishes in the pantry?",
            "zh": "茶水间清洗杯子/餐具，怎样更节水？",
        },
        "A": {
            "en": "Use a basin/filled sink for washing, then a quick rinse.",
            "zh": "先在盆/注水水槽中清洗，再快速冲洗。",
        },
        "B": {
            "en": "Wash under continuously running water until it looks clean.",
            "zh": "一直开着水龙头冲洗到看起来干净为止。",
        },
        "correct": "A",
        "explain": {
            "en": "Batch washing reduces continuous flow time and uses far less water.",
            "zh": "集中清洗能减少持续放水时间，明显更省水。",
        },
    },
    {
        "q": {
            "en": "(At home) Dual-flush toilet etiquette to save water?",
            "zh": "（家中）双档冲水马桶的节水使用方式是？",
        },
        "A": {
            "en": "Use half flush for liquid waste; full flush only when necessary.",
            "zh": "液体污物用半冲，必要时才使用全冲。",
        },
        "B": {
            "en": "Always use full flush to be safe.",
            "zh": "为保险起见总是用全冲。",
        },
        "correct": "A",
        "explain": {
            "en": "Half flush is designed to save water for typical uses.",
            "zh": "半冲就是为常见场景节水而设计。",
        },
    },
    {
        "q": {
            "en": "(At home) How should you wash vegetables/fruit to save water?",
            "zh": "（家中）清洗蔬菜水果时如何节水？",
        },
        "A": {
            "en": "Soak/agitate in a bowl first, then finish with a short rinse.",
            "zh": "先用盆浸泡/搅拌清洗，再短时间冲洗。",
        },
        "B": {
            "en": "Rinse under running water for a long time.",
            "zh": "长时间在流水下冲洗。",
        },
        "correct": "A",
        "explain": {
            "en": "Soaking does most of the cleaning with minimal water; a brief rinse completes the job.",
            "zh": "浸泡即可完成大部分清洁，最后短时间冲洗即可。",
        },
    },
    {
        "q": {
            "en": "(At home) How to clean a thermos/bottle interior without wasting water?",
            "zh": "（家中）如何清洗保温杯/水壶内壁更省水？",
        },
        "A": {
            "en": "Soak with detergent and brush, then rinse briefly.",
            "zh": "加入清洁剂浸泡并用刷子刷洗，最后短时间冲洗。",
        },
        "B": {
            "en": "Rinse under strong flow for an extended time.",
            "zh": "长时间在强水流下冲洗。",
        },
        "correct": "A",
        "explain": {
            "en": "Soaking + brushing is effective and uses much less water than extended rinsing.",
            "zh": "“浸泡+刷洗”更有效，比长时间冲洗省水得多。",
        },
    },
]

# Map each question to its corresponding illustrative photo.  The
# images have been cropped to a uniform 600×400 px size and are
# expected to reside alongside this script in the same directory.
IMAGE_FILES = [
    os.path.join(os.path.dirname(__file__), f"crop_q{i}.jpg") for i in range(1, 8)
]


class ABQuizApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.lang = "en"
        # Keep a reference to the current question dict
        self.current = None
        # Lock prevents answering before drawing a new question
        self.locked = True
        # Placeholder for PhotoImage to prevent garbage collection
        self.photo = None

        # Set window properties
        self.root.title(APP_TITLE_EN)
        self.root.geometry("900x600")

        self.build_ui()
        self.update_left()
        self.bind_hotkeys()

    def t(self, key: str) -> str:
        """Simple translation helper based on current language."""
        text = {
            "cards_left_en": "Total questions",
            "cards_left_zh": "题库总数",
            "draw_en": "Draw Question",
            "draw_zh": "抽题",
            "next_en": "Next",
            "next_zh": "下一题",
            "restart_en": "Restart",
            "restart_zh": "重开",
            "start_hint_en": "Click “Draw Question” to start.",
            "start_hint_zh": "点击“抽题”开始答题。",
            "feedback_correct_en": "✅ Correct!",
            "feedback_correct_zh": "✅ 回答正确！",
            "feedback_wrong_en": "❌ Not quite. Recommended answer is",
            "feedback_wrong_zh": "❌ 不完全正确。推荐答案是",
            "rec_answer_en": "Recommended answer",
            "rec_answer_zh": "推荐答案",
            "why_en": "Why",
            "why_zh": "原因",
            "lang_btn_en": "中文",
            "lang_btn_zh": "English",
            "leak_tip_en": f"If you see leaks or continuous running water, report to Facilities ({FACILITIES_HOTLINE_EN}).",
            "leak_tip_zh": f"若发现漏水或持续流水，请及时联系设施报修（{FACILITIES_HOTLINE_ZH}）。",
        }
        return text[f"{key}_{self.lang}"]

    def build_ui(self) -> None:
        """Construct all of the widgets used in the application."""
        # Top bar containing remaining count and control buttons
        top = tk.Frame(self.root)
        top.pack(fill="x", pady=(10, 6))

        self.lbl_left = tk.Label(top, text="", font=("Segoe UI", 12, "bold"))
        self.lbl_left.pack(side="left", padx=10)

        btns = tk.Frame(top)
        btns.pack(side="right")

        self.btn_lang = tk.Button(btns, text=self.t("lang_btn"), width=8, command=self.toggle_lang)
        self.btn_lang.pack(side="right", padx=(8, 0))
        self.btn_draw = tk.Button(btns, text=self.t("draw"), width=14, command=self.draw_question)
        self.btn_draw.pack(side="right", padx=4)
        # Next and restart simply draw a new question
        self.btn_next = tk.Button(btns, text=self.t("next"), width=10, command=self.draw_question)
        self.btn_next.pack(side="right", padx=4)
        self.btn_restart = tk.Button(btns, text=self.t("restart"), width=10, command=self.draw_question)
        self.btn_restart.pack(side="right", padx=4)

        # Main card area
        self.card = tk.Frame(self.root, bd=2, relief="ridge", bg="#ffffff")
        self.card.pack(expand=True, fill="both", padx=16, pady=10)

        # Question area (blue background)
        self.q_area = tk.Frame(self.card, bg="#d0e9f7")
        self.q_area.pack(fill="x", padx=0, pady=(0, 0))

        # Question label
        self.lbl_q = tk.Label(
            self.q_area,
            text=self.t("start_hint"),
            wraplength=820,
            justify="left",
            font=("Segoe UI", 17, "bold"),
            bg="#d0e9f7",
        )
        self.lbl_q.pack(padx=16, pady=(20, 12), anchor="w")

        # Options within question area
        opts = tk.Frame(self.q_area, bg="#d0e9f7")
        opts.pack(fill="x", padx=16, pady=(8, 12))
        self.btn_A = tk.Button(opts, text="A", font=("Segoe UI", 14, "bold"), width=3, command=lambda: self.answer("A"))
        self.btn_B = tk.Button(opts, text="B", font=("Segoe UI", 14, "bold"), width=3, command=lambda: self.answer("B"))
        self.lbl_A = tk.Label(opts, text="", wraplength=740, justify="left", font=("Segoe UI", 14), bg="#d0e9f7")
        self.lbl_B = tk.Label(opts, text="", wraplength=740, justify="left", font=("Segoe UI", 14), bg="#d0e9f7")
        # grid layout for options
        self.btn_A.grid(row=0, column=0, sticky="n")
        self.lbl_A.grid(row=0, column=1, sticky="w", padx=(10, 0))
        self.btn_B.grid(row=1, column=0, sticky="n", pady=(10, 0))
        self.lbl_B.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=(10, 0))

        # Display area for image and feedback/official answer
        self.display_area = tk.Frame(self.card, bg="#ffffff")
        self.display_area.pack(expand=True, fill="both", padx=16, pady=(0, 10))

        # Image label (photo shown before answering)
        self.img_label = tk.Label(self.display_area, bg="#ffffff")
        self.img_label.pack()

        # Feedback label for correct/incorrect text
        self.lbl_feedback = tk.Label(self.display_area, text="", font=("Segoe UI", 13, "bold"), bg="#ffffff")

        # Recommended answer and explanation label
        self.lbl_official = tk.Label(
            self.display_area,
            text="",
            wraplength=820,
            justify="left",
            font=("Segoe UI", 15, "bold"),
            fg="#1f4d1f",
            bg="#ffffff",
        )

        # Tip below card
        self.lbl_tip = tk.Label(
            self.root,
            text=self.t("leak_tip"),
            font=("Segoe UI", 10),
            fg="#666",
        )
        self.lbl_tip.pack(pady=(0, 8))

    def toggle_lang(self) -> None:
        """Switch between English and Chinese and refresh the UI labels."""
        self.lang = "zh" if self.lang == "en" else "en"
        self.root.title(APP_TITLE_ZH if self.lang == "zh" else APP_TITLE_EN)
        # Update button text
        self.btn_lang.config(text=self.t("lang_btn"))
        self.btn_draw.config(text=self.t("draw"))
        self.btn_next.config(text=self.t("next"))
        self.btn_restart.config(text=self.t("restart"))
        self.lbl_tip.config(text=self.t("leak_tip"))
        # Update remaining count
        self.update_left()
        # Refresh current question and answers
        if self.current:
            self.show_question(self.current)

    def update_left(self) -> None:
        """Update the label showing the total number of questions."""
        total = len(QUESTION_BANK)
        self.lbl_left.config(text=f"{self.t('cards_left')}: {total}")

    def draw_question(self) -> None:
        """Randomly select a question and display it along with its image."""
        self.current = random.choice(QUESTION_BANK)
        self.locked = False
        self.show_question(self.current)

    def show_question(self, q: dict) -> None:
        """Display the current question, options and associated image."""
        # Set question text and option texts in current language
        self.lbl_q.config(text=q["q"][self.lang])
        self.lbl_A.config(text=q["A"][self.lang])
        self.lbl_B.config(text=q["B"][self.lang])
        # Hide feedback and official answer labels if previously shown
        self.lbl_feedback.pack_forget()
        self.lbl_official.pack_forget()
        # Determine image index based on question order
        try:
            idx = QUESTION_BANK.index(q)
        except ValueError:
            idx = 0
        img_path = IMAGE_FILES[idx]
        # Load and display the image
        try:
            pil_img = Image.open(img_path)
            self.photo = ImageTk.PhotoImage(pil_img)
            self.img_label.config(image=self.photo)
            # Ensure image label is visible
            if not self.img_label.winfo_ismapped():
                self.img_label.pack()
        except Exception as e:
            # If image fails to load, hide the label
            self.img_label.config(image="")
        # Reset feedback text to empty
        self.lbl_feedback.config(text="")
        self.lbl_official.config(text="")

    def answer(self, choice: str) -> None:
        """Handle the player's answer and reveal the explanation."""
        if self.locked or not self.current:
            return
        correct = self.current["correct"]
        # Determine correct/incorrect feedback text and colours
        if choice == correct:
            feedback_text = self.t("feedback_correct")
            feedback_color = "#0a7a0a"
        else:
            feedback_text = f"{self.t('feedback_wrong')} {correct}."
            feedback_color = "#b00020"
        # Update feedback label
        self.lbl_feedback.config(text=feedback_text, fg=feedback_color)
        # Hide the image when answer is given
        self.img_label.pack_forget()
        # Compose recommended answer and explanation
        exp = self.current["explain"][self.lang]
        ans_text = self.current[correct][self.lang]
        official_text = f"{self.t('rec_answer')} ({correct}): {ans_text}\n{self.t('why')}: {exp}"
        self.lbl_official.config(text=official_text)
        # Show feedback and official answer labels
        self.lbl_feedback.pack(anchor="w", padx=16, pady=(10, 4))
        self.lbl_official.pack(anchor="w", padx=16, pady=(4, 10))
        # Lock until next question is drawn
        self.locked = True

    def bind_hotkeys(self) -> None:
        """Register keyboard shortcuts for quick answers and drawing new questions."""
        # Answer with A/B using lowercase or uppercase keys
        self.root.bind("<Key-a>", lambda e: self.answer("A"))
        self.root.bind("<Key-A>", lambda e: self.answer("A"))
        self.root.bind("<Key-b>", lambda e: self.answer("B"))
        self.root.bind("<Key-B>", lambda e: self.answer("B"))
        # Return draws the next question
        self.root.bind("<Return>", lambda e: self.draw_question())


if __name__ == "__main__":
    root = tk.Tk()
    app = ABQuizApp(root)
    root.mainloop()