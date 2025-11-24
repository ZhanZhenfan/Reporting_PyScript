"""
Modified Smart Saver Ranking Game
================================

This script is a heavily customised version of the original "Smart Saver Ranking
Game".  The original game asked the user to rank a handful of household or
office appliances by their power consumption.  In this version the ranking
logic, the card faces and backs, and the overall messaging have been
substantially altered to teach players about realistic, scenario‑based energy
waste.  Instead of assuming every appliance runs 24/7, each device is now
associated with a plausible misuse scenario and an estimated daily usage in
hours.  Annual electricity cost is computed from the device’s wattage, the
assumed hours per day and the local tariff, and this cost forms the basis of
the ranking.  When the cards are flipped over, players see not only the
annual cost but also a simple suggestion on how to save energy.

Key changes in this version include:

* **Scenario‑based usage assumptions.**  Each appliance is mapped to an
  estimated number of hours it might run per day if left on unnecessarily.
  These estimates are drawn from typical usage patterns and advice from
  reputable sources.  For example, a central air conditioner might run for
  6–8 hours in moderate climates but can exceed 12 hours in hot climates【308214399740575†L360-L363】; a
  refrigerator’s compressor cycles on for roughly 30 % of the day (about
  eight hours)【75863526599615†L104-L112】; and a projector left on after a meeting will
  operate for many hours【715541038231362†L114-L118】.  These assumptions replace the
  previous 24/7 calculation.

* **Revised equipment list.**  The coffee machine has been removed to avoid
  duplication and simplify the game.  Other appliances remain, each with its
  wattage and daily hours defined.

* **Updated ranking logic.**  The correct order is now based on the annual cost
  computed from `(watts/1000) * hours_per_day * 365 * tariff` instead of the
  instantaneous wattage.  This emphasises how long‑running, even low‑powered
  devices like lighting can add up over a year.

* **Front of the cards.**  In addition to the device name and an image, the
  front of each card now shows a short description of the misuse scenario
  together with the assumed hours per day (for example “Not turned off after
  meeting (12 h/day)” for a projector).  This helps players understand why the
  device appears in the ranking and what behaviour leads to waste.

* **Back of the cards.**  When flipped, the back still displays the device
  name and the annual cost but also includes the actual hours‑per‑day
  assumption and a concise suggestion for saving energy.  The cost line has
  been moved up slightly to make room for the suggestion at the bottom.

* **Simplified answer check.**  The popup that previously listed the correct
  order has been removed.  When the user clicks “Check Answer” the cards
  flash green or red and a sound plays, then the annual cost and advice are
  revealed.  No extra dialog appears because the back of each card now
  contains the relevant answer.

To run this script you will need Python’s Tkinter library and Pillow (for
loading JPEG images).  Place your JPEG files in the same directory as this
script, using the names defined in `IMAGE_FILENAMES`.  If an image is
missing, a placeholder will be drawn instead.
"""

import os
import random
import sys
import time
import tkinter as tk
from tkinter import messagebox

try:
    # Pillow is required for JPEG image loading.  If unavailable, the game
    # will still run but will draw placeholders instead of photos.
    from PIL import Image, ImageTk  # type: ignore
    PIL_OK = True
except Exception as e:
    PIL_OK = False
    PIL_ERR = e

# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------

# Singapore electricity tariff (SGD per kWh).  Updated for Q4 2025.
# According to SP Group’s regulated tariff for Q4 2025, the rate (before GST)
# increased from 27.47 ¢/kWh to 27.55 ¢/kWh【252054037803380†L320-L323】.  This corresponds to
# roughly S$0.2755 per kWh.  We use the before‑GST rate here because it is
# commonly quoted by SP Group and avoids uncertainty in GST changes.
SG_TARIFF_SGD_PER_KWH = 0.2755

# Equipment definitions.  Each entry maps an appliance name to a tuple of
# (wattage, hours_per_day).  The daily hours reflect a plausible misuse
# scenario; see the module docstring for discussion.  Coffee Machine has been
# removed.
EQUIPMENT: dict[str, tuple[int, float]] = {
    "Central Air Conditioner": (3500, 12.0),  # left on after work (12 h/day)
    "Microwave": (800, 1.0),                  # multiple heating sessions (1 h/day)
    "Water Dispenser (Heating)": (650, 4.0),  # overheats water due to excess use (4 h/day)
    "Refrigerator": (500, 8.0),              # compressor runs 30 % of the time (8 h/day)
    "Printer (Printing)": (400, 2.0),        # long print jobs (2 h/day)
    "Projector": (250, 12.0),                # not turned off after meeting (12 h/day)
    "LED Tube": (18, 12.0),                  # lights left on when room is empty (12 h/day)
    "Desktop PC": (150, 8.0),                # idle or unused computer (8 h/day)
    "Laptop": (45, 8.0),                    # charging/standby for hours (8 h/day)
}

# Scenario descriptions for the front of each card.  Each entry provides a short
# English sentence describing the misuse scenario and includes the hours per
# day.  These strings appear below the device name on the card front.
SCENARIOS: dict[str, str] = {
    "Central Air Conditioner": "Left on after work (12 h/day)",
    "Microwave": "Multiple heating sessions (1 h/day)",
    "Water Dispenser (Heating)": "Overheating due to excess use (4 h/day)",
    "Refrigerator": "Compressor runs to keep cold (8 h/day)",
    "Printer (Printing)": "Long printing tasks (2 h/day)",
    "Projector": "Not turned off after meeting (12 h/day)",
    "LED Tube": "Lights left on when room is empty (12 h/day)",
    "Desktop PC": "Idle or unused for hours (8 h/day)",
    "Laptop": "Charging/standby for hours (8 h/day)",
}

# Energy‑saving suggestions for the back of each card.  These messages appear
# beneath the cost to help players understand how to reduce waste.
SUGGESTIONS: dict[str, str] = {
    "Central Air Conditioner": "Turn off or adjust thermostat when you leave.",
    "Microwave": "Use only when necessary and unplug when not in use.",
    "Water Dispenser (Heating)": "Heat only the water you need and turn off heating.",
    "Refrigerator": "Minimise door opening and maintain proper temperature.",
    "Printer (Printing)": "Print only when necessary and turn off afterwards.",
    "Projector": "Turn off immediately after the meeting or set an auto‑off timer.",
    "LED Tube": "Turn off lights when you leave or install motion sensors.",
    "Desktop PC": "Shut down or use sleep mode when idle.",
    "Laptop": "Unplug charger when full and power down when not in use.",
}

# Map each device name to the expected JPEG file name.  Place these files in
# the same directory as this script.  If an image is missing, a placeholder
# graphic will be drawn instead.  You may customise the filenames as long as
# they match your assets.
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

# Determine the absolute path of this script to locate images relative to it.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGE_PATHS = {name: os.path.join(BASE_DIR, fname) for name, fname in IMAGE_FILENAMES.items()}


class DragDropGame:
    """
    A drag‑and‑drop ranking game that teaches users about the real cost of
    leaving appliances on unnecessarily.  Players must drag the cards into the
    correct order from highest to lowest annual electricity cost.  Each card
    describes a misuse scenario and the assumed hours per day that the device
    operates.  After the user checks the order, the annual cost and a simple
    energy‑saving suggestion are shown on the back of each card.
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Smart Saver Ranking Game (Highest → Lowest Cost)")
        # A larger window to accommodate big images
        self.root.geometry("1400x720")

        # Number of cards per round
        self.num_cards: int = 4  # reduce to four cards instead of five

        # Game state
        self.sample: list[tuple[str, tuple[int, float]]] = []      # Selected devices for current round
        self.correct_order: list[str] = []           # Names sorted by annual cost descending
        self.slot_centers: list[tuple[float, float]] = []  # X,Y centers of slots
        self.card_for_slot: dict[int, str] = {}      # Slot index → card tag
        self.slot_for_card: dict[str, int] = {}      # Card tag → slot index
        # card tag → dict with rectangle id, list of image ids, list of text ids
        self.cards: dict[str, dict[str, list[int] | int]] = {}
        self.dragging_tag: str | None = None
        self.drag_prev: tuple[int, int] = (0, 0)
        self.last_motion_ts = 0.0
        self.already_revealed: bool = False          # True when answers shown
        self.image_cache: dict[str, ImageTk.PhotoImage] = {}


        # Layout and sizing.  The card dimensions have been increased to
        # provide more space for larger images and text.  Feel free to
        # customise these values further to suit your images and design.
        self.front_img_max_w = 280  # wider image area
        self.front_img_max_h = 360  # taller image area
        self.left_margin = 30
        self.top_margin = 240
        # Slot dimensions: card width equals image width plus padding, card
        # height equals image height plus extra space for text.  Increased
        # values accommodate larger fonts and suggestions on the back.
        self.slot_w = self.front_img_max_w + 60
        self.slot_h = self.front_img_max_h + 180
        self.slot_gap = 18
        self.card_pad = 10
        self.anim_steps = 8
        self.anim_delay = 12

        # Fonts for different parts of the card
        # Font definitions.  All font sizes have been increased slightly to
        # improve legibility on the larger cards.
        self.font_front_name = ("Segoe UI", 14, "bold")
        # Scenario text is bold because it explains the assumption behind the
        # annual cost calculation; this is an important educational element.
        self.font_front_scenario = ("Segoe UI", 13, "bold")
        self.font_back_title = ("Segoe UI", 14, "bold")
        self.font_back_cost = ("Segoe UI", 22, "bold")
        self.font_back_hint = ("Segoe UI", 12)
        # Suggestions are displayed boldly and slightly larger to emphasise
        # recommended actions; these are critical to the learning outcome.
        self.font_back_suggestion = ("Segoe UI", 14, "bold")

        # GUI header
        header = tk.Label(
            root,
            # Three arrows correspond to four cards (Highest → → → Lowest)
            text="Drag to rank  (Left → Right)  Highest  →  →  →  Lowest annual cost",
            font=("Segoe UI", 16, "bold")
        )
        header.pack(pady=(10, 6))

        # Top bar buttons
        tb = tk.Frame(root)
        tb.pack()
        # Button now draws four cards instead of five
        tk.Button(tb, text="Draw 4 Cards", command=self.start_game).pack(side="left", padx=6)
        tk.Button(tb, text="Check Answer", command=self.on_check_and_reveal).pack(side="left", padx=6)
        tk.Button(tb, text="Reset", command=self.reset_board).pack(side="left", padx=6)

        # Canvas for drawing cards and slots
        self.canvas = tk.Canvas(root, bg="#f7f8fb", highlightthickness=0)
        self.canvas.pack(expand=True, fill="both", padx=16, pady=10)

        # Informational text on the canvas
        self.hint_id = self.canvas.create_text(
            0, 0,
            text="Place cards Left → Right  (Highest → Lowest annual cost)",
            font=("Segoe UI", 12), fill="#666"
        )
        self.tariff_id = self.canvas.create_text(
            0, 0,
            text=(
                f"Annual cost = watts × hours/day × 365 × tariff  •  Tariff: S${SG_TARIFF_SGD_PER_KWH:.3f}/kWh"
            ),
            font=("Segoe UI", 10), fill="#888"
        )

        # Bind resize and drag events
        self.canvas.bind("<Configure>", self.on_resize)
        self.canvas.tag_bind("card", "<ButtonPress-1>", self.on_press)
        self.canvas.tag_bind("card", "<B1-Motion>", self.on_drag)
        self.canvas.tag_bind("card", "<ButtonRelease-1>", self.on_release)

        # Check availability of images and notify in console
        self._log_missing_images()
        if not PIL_OK:
            messagebox.showwarning(
                "Pillow missing",
                "Pillow is required for JPG images.\nInstall with:\n\npip install pillow\n\n" + str(PIL_ERR)
            )

    # -------------------------------------------------------------------------
    # Utility methods
    # -------------------------------------------------------------------------
    def _log_missing_images(self) -> None:
        """Print missing image files to the console for debugging."""
        missing = [(n, os.path.basename(p)) for n, p in IMAGE_PATHS.items() if not os.path.exists(p)]
        if missing:
            print(">> Missing image files (place these .jpg in the same folder):")
            for n, fname in missing:
                print(f"   - {n}  ->  {fname}")
        else:
            print(">> All standard JPG filenames found.")

    def reset_board(self) -> None:
        """Clear the canvas and reset all game state."""
        self.canvas.delete("all")
        self.hint_id = self.canvas.create_text(
            0, 0,
            text="Place cards Left → Right  (Highest → Lowest annual cost)",
            font=("Segoe UI", 12), fill="#666"
        )
        self.tariff_id = self.canvas.create_text(
            0, 0,
            text=(
                f"Annual cost = watts × hours/day × 365 × tariff  •  Tariff: S${SG_TARIFF_SGD_PER_KWH:.4f}/kWh (Q4 2025)"
            ),
            # Bold and larger font to emphasise the tariff assumption
            font=("Segoe UI", 11, "bold"), fill="#666"
        )
        self.sample.clear()
        self.correct_order.clear()
        self.slot_centers.clear()
        self.card_for_slot.clear()
        self.slot_for_card.clear()
        self.cards.clear()
        self.already_revealed = False

    def start_game(self) -> None:
        """Initialize a new round by drawing a fixed number of random equipment cards.

        In this version the number of cards drawn is controlled by
        ``self.num_cards``.  We default to four cards per round to keep
        the game concise and focused on the most impactful scenarios.
        """
        self.reset_board()
        # Randomly select a fixed number of devices
        self.sample = random.sample(list(EQUIPMENT.items()), self.num_cards)
        # Sort the selected devices by annual cost descending to get the correct order
        cost_list: list[tuple[str, float]] = []
        for name, (watts, hours) in self.sample:
            kwh_year = (watts / 1000.0) * (hours * 365)
            cost_year = kwh_year * SG_TARIFF_SGD_PER_KWH
            cost_list.append((name, cost_year))
        # Determine correct order based on descending annual cost
        self.correct_order = [n for n, _ in sorted(cost_list, key=lambda x: x[1], reverse=True)]
        self.layout_slots_and_cards()

    # -------------------------------------------------------------------------
    # Layout
    # -------------------------------------------------------------------------
    def on_resize(self, event: tk.Event) -> None:
        """Repaint the board on window resize, preserving card order and face."""
        if not self.sample:
            self.center_hint_only()
            return
        order_names: list[str | None] = []
        for i in range(self.num_cards):
            tag = self.card_for_slot.get(i)
            order_names.append(tag.replace("card_name_", "") if tag else None)
        self.layout_slots_and_cards(keep_order=order_names, keep_face=self.already_revealed)

    def center_hint_only(self) -> None:
        """Center the hint text when no cards are present."""
        w = self.canvas.winfo_width()
        self.canvas.coords(self.hint_id, w / 2, 90)
        self.canvas.coords(self.tariff_id, w / 2, 115)

    def layout_slots_and_cards(self, keep_order: list[str | None] | None = None,
                               keep_face: bool = False) -> None:
        """Lay out the slot rectangles and cards.

        If `keep_order` is provided, the cards will be arranged in that order
        (filling in any missing cards from the current sample).  If
        `keep_face` is True, the cards will remain on the face they were
        previously on (front or back).
        """
        self.canvas.delete("all")
        w = self.canvas.winfo_width()
        # Draw explanatory text
        self.hint_id = self.canvas.create_text(
            w / 2, 90,
            text="Left → Right = Highest → Lowest annual cost",
            font=("Segoe UI", 12), fill="#666"
        )
        self.tariff_id = self.canvas.create_text(
            w / 2, 115,
            text=(
                f"Annual cost = watts × hours/day × 365 × tariff  •  Tariff: S${SG_TARIFF_SGD_PER_KWH:.4f}/kWh (Q4 2025)"
            ),
            # Bold and larger font to emphasise the tariff assumption
            font=("Segoe UI", 11, "bold"), fill="#666"
        )

        # Determine slot positions based on window width
        total_min = self.num_cards * self.slot_w + (self.num_cards - 1) * self.slot_gap
        usable_w = max(total_min, min(w - 2 * self.left_margin, 2000))
        gap = self.slot_gap
        if usable_w > total_min and self.num_cards > 1:
            # Distribute extra space evenly across the gaps (one less than the number of cards)
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

        # Determine which devices to display and in what order
        if keep_order and any(keep_order):
            # Recreate the list of devices in the preserved order, filling
            # any missing slots with the remaining sampled devices.
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
            # Draw card background
            rect = self.canvas.create_rectangle(
                cx - half_w, cy - half_h, cx + half_w, cy + half_h,
                fill="#ffffff", outline="#5b8def", width=2,
                tags=("card", f"card_name_{name}")
            )
            # Draw image (or placeholder).  The image is raised slightly
            # above centre to make room for the scenario text below.
            img_ids = self.create_front_image(name, cx, cy - 50)
            # Draw text for front (device name)
            txt_y = cy + (self.front_img_max_h / 2) + 26
            txt_name = self.canvas.create_text(
                cx, txt_y, text=name,
                font=self.font_front_name, tags=("card", f"card_name_{name}")
            )
            # Scenario line just below the device name
            scenario_text = SCENARIOS.get(name, "")
            txt_scenario = self.canvas.create_text(
                cx, txt_y + 22, text=scenario_text,
                font=self.font_front_scenario, fill="#555", tags=("card", f"card_name_{name}")
            )
            # Store card elements
            self.cards[f"card_name_{name}"] = {
                "rect": rect,
                "img_ids": img_ids,
                "txt_ids": [txt_name, txt_scenario],
            }
            self.card_for_slot[i] = f"card_name_{name}"
            self.slot_for_card[f"card_name_{name}"] = i
            # If keeping face (back), reveal cost instead of name
            if keep_face:
                self.show_back(f"card_name_{name}", name)

        # Bind drag events to the newly created cards
        self.canvas.tag_bind("card", "<ButtonPress-1>", self.on_press)
        self.canvas.tag_bind("card", "<B1-Motion>", self.on_drag)
        self.canvas.tag_bind("card", "<ButtonRelease-1>", self.on_release)

    # -------------------------------------------------------------------------
    # Image loading
    # -------------------------------------------------------------------------
    def load_photo(self, path: str, max_w: int, max_h: int) -> ImageTk.PhotoImage | None:
        """Load an image using Pillow and resize it to fit within max_w x max_h.

        Returns an ImageTk.PhotoImage object on success, or None if Pillow is
        unavailable, the file does not exist, or an error occurs.
        """
        if not PIL_OK or not path or not os.path.isfile(path):
            return None
        try:
            im = Image.open(path)  # type: ignore[attr-defined]
            im = im.convert("RGBA")  # Ensure consistent alpha channel
            im.thumbnail((max_w, max_h), Image.LANCZOS)
            return ImageTk.PhotoImage(im)  # type: ignore[attr-defined]
        except Exception as e:
            print(f">> Failed to load image: {path}\n   {e}")
            return None

    def create_front_image(self, name: str, cx: float, cy: float) -> list[int]:
        """Create an image on the canvas for the given device.

        If the image file cannot be loaded, draws a placeholder (a colored
        rectangle and an emoji).  Returns a list of canvas item IDs for the
        image (or placeholder elements) so they can be managed later.
        """
        ids: list[int] = []
        path = IMAGE_PATHS.get(name)
        img_obj = self.load_photo(path, self.front_img_max_w, self.front_img_max_h)
        if img_obj is not None:
            self.image_cache[name] = img_obj  # Keep reference to avoid garbage collection
            ids.append(self.canvas.create_image(
                cx, cy, image=img_obj, tags=("card", f"card_name_{name}")
            ))
        else:
            # Draw a placeholder box with an image icon
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

    # -------------------------------------------------------------------------
    # Sound feedback
    # -------------------------------------------------------------------------
    def play_correct_sound(self) -> None:
        """Play a short sound to indicate a correct answer.

        On Windows, uses winsound.Beep with a high tone; otherwise falls back
        to Tkinter's bell method.  If both methods fail, no sound is played.
        """
        try:
            if sys.platform.startswith("win"):
                import winsound  # type: ignore
                winsound.Beep(1000, 200)  # frequency 1kHz, duration 200 ms
            else:
                # Tkinter bell triggers the system default beep (if available)
                self.root.bell()
        except Exception:
            # Ignore sound errors on unsupported systems
            pass

    def play_wrong_sound(self) -> None:
        """Play a short sound to indicate an incorrect answer.

        Uses a lower tone than the success sound on Windows, or the default
        bell on other platforms.
        """
        try:
            if sys.platform.startswith("win"):
                import winsound  # type: ignore
                winsound.Beep(400, 500)  # frequency 400 Hz, duration 500 ms
            else:
                self.root.bell()
        except Exception:
            pass

    def flash_cards(self, correct: bool) -> None:
        """Flash the borders of all cards briefly in green or red.

        Called after the user checks the order.  Borders turn green for a
        correct arrangement and red for an incorrect one, then revert to
        their default color after a short delay.
        """
        color = "#28a745" if correct else "#dc3545"
        for tag in self.card_for_slot.values():
            rect_id = self.cards[tag]["rect"]  # type: ignore[index]
            self.canvas.itemconfig(rect_id, outline=color, width=3)

        def revert() -> None:
            for tag in self.card_for_slot.values():
                rect_id = self.cards[tag]["rect"]  # type: ignore[index]
                self.canvas.itemconfig(rect_id, outline="#5b8def", width=2)
        self.canvas.after(300, revert)

    # -------------------------------------------------------------------------
    # Drag handling
    # -------------------------------------------------------------------------
    def on_press(self, event: tk.Event) -> None:
        """Start dragging a card when the user presses the mouse button."""
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
        """Move the card as the user drags the mouse."""
        if not self.dragging_tag:
            return
        now = time.time()
        # Throttle dragging events slightly for smoother movement
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
        """Place the card into the nearest slot on mouse release."""
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
                # Find the first empty slot among the available slots
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

    # -------------------------------------------------------------------------
    # Helper methods for dragging
    # -------------------------------------------------------------------------
    def card_center(self, tag: str) -> tuple[float, float]:
        """Calculate the center of the card's rectangle."""
        rect_id = self.cards[tag]["rect"]  # type: ignore[index]
        x1, y1, x2, y2 = self.canvas.coords(rect_id)
        return ((x1 + x2) / 2, (y1 + y2) / 2)

    def animate_to(self, tag: str, target_xy: tuple[float, float]) -> None:
        """Animate the card to the target slot center."""
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
        """Move a card by a specific delta (without throttling)."""
        card = self.cards[tag]
        self.canvas.move(card["rect"], dx, dy)
        for iid in card["img_ids"]:  # type: ignore[index]
            self.canvas.move(iid, dx, dy)
        for tid in card["txt_ids"]:  # type: ignore[index]
            self.canvas.move(tid, dx, dy)

    def nearest_slot(self, x: float, y: float) -> int:
        """Find the slot index whose center is closest to (x, y)."""
        best = 0
        best_d = float("inf")
        for i, (sx, sy) in enumerate(self.slot_centers):
            d = (sx - x) ** 2 + (sy - y) ** 2
            if d < best_d:
                best_d = d
                best = i
        return best

    # -------------------------------------------------------------------------
    # Checking answers and revealing results
    # -------------------------------------------------------------------------
    def on_check_and_reveal(self) -> None:
        """Check the order of the cards and reveal the annual cost and advice."""
        if not self.sample:
            return
        names: list[str] = []
        for i in range(self.num_cards):
            tag = self.card_for_slot.get(i)
            if not tag:
                messagebox.showinfo("Result", "Please place all cards into slots.")
                return
            names.append(tag.replace("card_name_", ""))

        is_correct = names == self.correct_order
        if is_correct:
            self.play_correct_sound()
        else:
            self.play_wrong_sound()
        # Visual flash of card borders
        self.flash_cards(is_correct)
        # Reveal cost values if not done already
        if not self.already_revealed:
            for i in range(self.num_cards):
                tag = self.card_for_slot[i]
                name = tag.replace("card_name_", "")
                self.show_back(tag, name)
            self.already_revealed = True
        else:
            messagebox.showinfo(
                "Shown", "Answers are already displayed. Use Reset to start a new round."
            )

    def show_back(self, tag: str, name: str) -> None:
        """Reveal the annual cost and suggestion on the card.

        Deletes the front image and text, draws the cost in large font, and
        draws a small hint below along with a suggestion.  The text is
        centered and prominently displayed.  This method is called for each
        card when the answers are revealed.
        """
        watts, hours = EQUIPMENT[name]
        kwh_year = (watts / 1000.0) * (hours * 365)
        cost_year = kwh_year * SG_TARIFF_SGD_PER_KWH
        items = self.cards[tag]
        rect_id, img_ids, txt_ids = items["rect"], items["img_ids"], items["txt_ids"]
        # Delete existing images
        for iid in img_ids:  # type: ignore[index]
            try:
                self.canvas.delete(iid)
            except Exception:
                pass
        items["img_ids"] = []  # no images on back
        # Remove all text items (device name and scenario)
        for tid in txt_ids:  # type: ignore[index]
            try:
                self.canvas.delete(tid)
            except Exception:
                pass
        items["txt_ids"] = []
        # Change card background and outline
        self.canvas.itemconfig(rect_id, fill="#f2f5ff", outline="#3a6be0", width=2)
        # Compute center coordinates
        x1, y1, x2, y2 = self.canvas.coords(rect_id)
        cx = (x1 + x2) / 2
        cy = (y1 + y2) / 2
        # Title (device name) positioned higher up
        title_id = self.canvas.create_text(
            cx, cy - 60, text=name, font=self.font_back_title,
            fill="#0b3d91", tags=("card", tag)
        )
        # Large cost (annual cost).  Move this further up so it doesn't crowd
        # the hint and suggestion.  A little extra space above helps separate
        # the pricing line from the hint below.
        cost_str = f"≈ S${cost_year:,.2f} / year"
        cost_id = self.canvas.create_text(
            cx, cy - 40, text=cost_str, font=self.font_back_cost,
            fill="#111", tags=("card", tag)
        )
        # Hint showing hours per day and days per year.  This sits just
        # beneath the cost line, leaving room above and below.
        hint_str = f"({hours:.0f}\u00a0h/day × 365\u00a0d)"
        hint_id = self.canvas.create_text(
            cx, cy + 10, text=hint_str, font=self.font_back_hint,
            fill="#777", tags=("card", tag)
        )
        # Suggestion text.  Place it even lower on the card to create a clear
        # separation from the hint.  A larger vertical gap emphasises the
        # instructional nature of this line.
        base_suggestion = SUGGESTIONS.get(name, "")
        suggestion_text = f"Suggested action: {base_suggestion}"
        wrap_width = self.slot_w - 2 * self.card_pad - 20
        suggestion_id = self.canvas.create_text(
            cx, cy + 90, text=suggestion_text, font=self.font_back_suggestion,
            fill="#444", width=wrap_width, tags=("card", tag)
        )
        # Update the list of text ids
        items["txt_ids"] = [title_id, cost_id, hint_id, suggestion_id]


if __name__ == "__main__":
    root_app = tk.Tk()
    game = DragDropGame(root_app)
    root_app.mainloop()