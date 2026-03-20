"""
makita_xgt.py
Makita XGT 40V module for OBI-1 with unified Raspberry Pi Pico firmware.

Wiring:
• LXT:  Battery Data → Pico GP6 + 4.7k pull-up
• XGT:  Battery TR → 2N3904 collector + 4.7k pull-up
         2N3904 base (via 1k) → Pico GP4
         Pico GP5 → Battery TR

Mode switching is automatic:
• LXT module sends b'\xFF\x00'
• XGT module sends b'\xFF\x01'

No re-flashing needed — just switch modules in the GUI.

Uses direct raw serial after mode switch (b'\xFF\x01') because XGT protocol is different from LXT 1-Wire.

"""

import time
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QPushButton, QTreeWidget, QTreeWidgetItem,
    QLabel, QMessageBox, QHeaderView, QFrame
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush

# =============================================================================
# XGT raw commands (exact match to m5din-makita-xgt.ino)
# =============================================================================
MODEL_CMD         = bytes([0xA5, 0xA5, 0x00, 0x58, 0x0A, 0xD4, 0xB2, 0x32, 0x00, 0xD3, 0xC8, 0xE0, 0x00, 0x60, 0x00, 0xC0, 0x00, 0x80, 0xC8, 0xD0, 0x40, 0xDC, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF])
NUM_CHARGES_CMD   = bytes([0x33, 0xC8, 0x03, 0x00, 0x2A, 0x00, 0x00, 0xCC])
CELL_SIZE_CMD     = bytes([0x33, 0x27, 0xBB, 0x10, 0x00, 0x00, 0x00, 0xCC])
PARALLEL_CNT_CMD  = bytes([0x33, 0x67, 0xBB, 0x50, 0x00, 0x00, 0x00, 0xCC])
BATT_HEALTH_CMD   = bytes([0x33, 0xC4, 0x03, 0x00, 0x26, 0x00, 0x00, 0xCC])
CHARGE_CMD        = bytes([0x33, 0x13, 0x03, 0x80, 0x10, 0x00, 0x00, 0xCC])
TEMP1_CMD         = bytes([0x33, 0x3B, 0x03, 0xC0, 0x58, 0x00, 0x00, 0xCC])
TEMP2_CMD         = bytes([0x33, 0x7B, 0x03, 0xC0, 0x38, 0x00, 0x00, 0xCC])
PACK_VOLTAGE_CMD  = bytes([0x33, 0x43, 0x03, 0xC0, 0x00, 0x00, 0x00, 0xCC])
LOCK_STATUS_CMD   = bytes([0x33, 0xF8, 0x03, 0x00, 0x06, 0x00, 0x00, 0xCC])
RESET_CMD         = bytes([0x33, 0xC8, 0x9B, 0x69, 0xA5, 0x00, 0x00, 0xCC])
RESET_CMD1        = bytes([0x33, 0x00, 0x4B, 0xF4, 0x00, 0x00, 0x00, 0xCC])

# Bit-reverse table (exact same as .ino)
LOOKUP = [0x0, 0x8, 0x4, 0xc, 0x2, 0xa, 0x6, 0xe, 0x1, 0x9, 0x5, 0xd, 0x3, 0xb, 0x7, 0xf]

# Same display keys as makita_lxt.py for identical tree structure
INITIAL_DATA: dict[str, str] = {
    "Model":                    "",
    "Battery Type":             "XGT",
    "Capacity":                 "",
    "Charge count*":            "",
    "State":                    "",
    "Status code":              "N/A",
    "Manufacturing date":       "N/A",
    "State of Charge":          "",
    "Health":                   "",
    "Pack Voltage":             "",
    "Cell 1 Voltage":           "",
    "Cell 2 Voltage":           "",
    "Cell 3 Voltage":           "",
    "Cell 4 Voltage":           "",
    "Cell 5 Voltage":           "",
    "Cell 6 Voltage":           "",
    "Cell 7 Voltage":           "",
    "Cell 8 Voltage":           "",
    "Cell 9 Voltage":           "",
    "Cell 10 Voltage":          "",
    "Cell Voltage Difference":  "",
    "Temperature Sensor 1":     "",
    "Temperature Sensor 2":     "",
    "Overdischarge count":      "N/A",
    "Overdischarge %":          "N/A",
    "Overload count":           "N/A",
    "Overload %":               "N/A",
    "ROM ID":                   "N/A",
    "Battery message":          "N/A",
}

ROW_EVEN = QColor("#1C1F26")
ROW_ODD  = QColor("#20242C")


def get_display_name() -> str:
    return "Makita XGT"


class ModuleApplication(QWidget):
    def __init__(self, parent=None, interface_module=None, obi_instance=None):
        super().__init__(parent)
        self.interface = None
        self.obi_instance = obi_instance
        self._build_ui()
        self._insert_battery_data(INITIAL_DATA.copy())

    # ── public API (same as LXT) ─────────────────────────────────────────────
    def set_interface(self, interface_instance):
        self.interface = interface_instance
        if hasattr(interface_instance, 'ready'):
            interface_instance.ready.connect(self._on_interface_ready)
        if hasattr(interface_instance, 'disconnected'):
            interface_instance.disconnected.connect(self._on_interface_disconnected)

    # ── UI (exact copy of LXT layout) ────────────────────────────────────────
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        title = QLabel(get_display_name())
        title.setStyleSheet("font-size: 18pt; font-weight: 700; color: #00B4CC; letter-spacing: 1px;")
        title.setAlignment(Qt.AlignCenter)

        subtitle = QLabel("40V XGT Battery Diagnostics  ·  Pico unified firmware")
        subtitle.setStyleSheet("font-size: 9pt; color: #3E4555; letter-spacing: 1px;")
        subtitle.setAlignment(Qt.AlignCenter)

        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setStyleSheet("color: #2E3340; margin: 2px 0;")

        root.addWidget(title)
        root.addWidget(subtitle)
        root.addWidget(divider)
        root.addWidget(self._build_button_row())
        root.addWidget(self._build_tree(), stretch=1)
        root.addLayout(self._build_bottom_bar())

    def _build_button_row(self) -> QWidget:
        container = QWidget()
        grid = QGridLayout(container)
        grid.setSpacing(10)

        # Read Data
        rg = QGroupBox("Read Data")
        rl = QVBoxLayout(rg)
        self.btn_read_model = QPushButton("Read Battery Model")
        self.btn_read_data = QPushButton("Read Battery Data")
        self.btn_read_model.setEnabled(False)
        self.btn_read_data.setEnabled(False)
        rl.addWidget(self.btn_read_model)
        rl.addWidget(self.btn_read_data)
        self.btn_read_model.clicked.connect(self._on_read_static_click)
        self.btn_read_data.clicked.connect(self._on_read_data_click)

        # Function Test (LEDs not supported on XGT)
        fg = QGroupBox("Function Test")
        fl = QVBoxLayout(fg)
        self.btn_leds_on = QPushButton("LED Test ON")
        self.btn_leds_off = QPushButton("LED Test OFF")
        self.btn_leds_on.setEnabled(False)
        self.btn_leds_off.setEnabled(False)
        fl.addWidget(self.btn_leds_on)
        fl.addWidget(self.btn_leds_off)
        self.btn_leds_on.clicked.connect(lambda: QMessageBox.information(self, "Info", "LED test not supported on XGT"))
        self.btn_leds_off.clicked.connect(lambda: QMessageBox.information(self, "Info", "LED test not supported on XGT"))

        # Reset Battery
        rsg = QGroupBox("Reset Battery")
        rsl = QVBoxLayout(rsg)
        self.btn_clear_errors = QPushButton("Clear Errors")
        self.btn_reset_message = QPushButton("Lockout Reset")
        self.btn_clear_errors.setEnabled(False)
        self.btn_reset_message.setEnabled(False)
        rsl.addWidget(self.btn_clear_errors)
        rsl.addWidget(self.btn_reset_message)
        self.btn_clear_errors.clicked.connect(lambda: QMessageBox.information(self, "Info", "Clear Errors not applicable on XGT"))
        self.btn_reset_message.clicked.connect(self._reset_battery)

        grid.addWidget(rg, 0, 0)
        grid.addWidget(fg, 0, 1)
        grid.addWidget(rsg, 0, 2)
        return container

    def _build_tree(self) -> QTreeWidget:
        self.tree = QTreeWidget()
        self.tree.setColumnCount(2)
        self.tree.setHeaderLabels(["Parameter", "Value"])
        self.tree.header().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tree.header().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tree.setAlternatingRowColors(False)
        self.tree.setRootIsDecorated(False)
        return self.tree

    def _build_bottom_bar(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        note = QLabel("* Charge count approximate  ·  Select rows then Copy to export")
        note.setStyleSheet("font-size: 8pt; color: #3E4555;")
        copy_btn = QPushButton("Copy Selected")
        clear_btn = QPushButton("Clear")
        copy_btn.clicked.connect(self._copy_to_clipboard)
        clear_btn.clicked.connect(self._clear_data)
        layout.addWidget(note, stretch=1)
        layout.addWidget(copy_btn)
        layout.addWidget(clear_btn)
        return layout

    # ── helpers (exact same as LXT) ──────────────────────────────────────────
    def _on_interface_ready(self):
        self._switch_to_xgt_mode()
        self.btn_read_model.setEnabled(True)
        self.btn_read_data.setEnabled(True)
        self.btn_clear_errors.setEnabled(True)
        self.btn_reset_message.setEnabled(True)
        self._log("[XGT] Interface ready – switched to XGT raw mode")

    def _on_interface_disconnected(self):
        self.btn_read_model.setEnabled(False)
        self.btn_read_data.setEnabled(False)
        self.btn_clear_errors.setEnabled(False)
        self.btn_reset_message.setEnabled(False)

    def _require_interface(self) -> bool:
        if not self.interface or not self.interface.serial.is_open:
            QMessageBox.critical(self, "No Interface", "Please connect the Arduino OBI interface first.")
            return False
        return True

    def _log(self, msg: str):
        if self.obi_instance:
            self.obi_instance.update_debug(msg)

    def _switch_to_xgt_mode(self):
        """Send mode-switch command to unified Pico firmware."""
        try:
            self.interface.serial.write(b'\xFF\x01')
            self.interface.serial.flush()
            time.sleep(0.1)
        except Exception as e:
            self._log(f"[XGT] Mode switch failed: {e}")

    # ── direct raw XGT communication (after mode switch) ─────────────────────
    def _bit_reverse(self, data: bytes) -> bytes:
        result = bytearray()
        for b in data:
            high = LOOKUP[b & 0x0F] << 4
            low = LOOKUP[b >> 4]
            result.append(high | low)
        return bytes(result)

    def _check_crc(self, buf: bytes) -> bool:
        if len(buf) < 8:
            return False
        if buf[0] == 0xCC:
            crc = buf[0]
            for b in buf[2:]:
                crc = (crc + b) % 256
            return crc == buf[1]
        else:
            length = len(buf) - ((buf[3] & 0x0F) + 2)
            crc = 0
            for b in buf[2:length]:
                crc += b
            return (buf[length] << 8 | buf[length + 1]) == crc

    def _send_xgt_command(self, cmd: bytes) -> bytes | None:
        """Raw serial send/receive for XGT (used after mode switch)."""
        try:
            self.interface.serial.reset_input_buffer()
            self.interface.serial.write(cmd)
            self.interface.serial.flush()
            time.sleep(0.05)
            resp = self.interface.serial.read(32)
            if len(resp) < 8:
                return None
            resp = self._bit_reverse(resp)
            if not self._check_crc(resp):
                return None
            return resp
        except Exception as e:
            self._log(f"[XGT] Command failed: {e}")
            return None

    # ── read functions ───────────────────────────────────────────────────────
    def _read_full_battery(self):
        if not self._require_interface():
            return

        try:
            self.interface.serial.write(b'\x00')  # wake-up
            time.sleep(0.07)

            data = INITIAL_DATA.copy()

            # Model
            resp = self._send_xgt_command(MODEL_CMD)
            if resp and len(resp) > 10:
                model_bytes = resp[-8:][::-1]
                data["Model"] = "".join(chr(b) for b in model_bytes if 32 <= b <= 126).strip()

            # Basic values
            cmds = [
                (NUM_CHARGES_CMD,   "Charge count*", lambda b: int.from_bytes(b[4:6], 'big')),
                (CELL_SIZE_CMD,     "Capacity",      lambda b: b[5] * 100),   # will multiply by parallel later
                (PARALLEL_CNT_CMD,  "Capacity",      lambda b: b[4]),         # multiplier
                (BATT_HEALTH_CMD,   "Health",        lambda b: int.from_bytes(b[4:6], 'big')),
                (CHARGE_CMD,        "State of Charge", lambda b: f"{int.from_bytes(b[4:6], 'big') // 255}%"),
                (TEMP1_CMD,         "Temperature Sensor 1", lambda b: f"{-30 + (int.from_bytes(b[4:6], 'big') - 2431) / 10:.1f} °C"),
                (TEMP2_CMD,         "Temperature Sensor 2", lambda b: f"{-30 + (int.from_bytes(b[4:6], 'big') - 2431) / 10:.1f} °C"),
                (PACK_VOLTAGE_CMD,  "Pack Voltage",  lambda b: f"{int.from_bytes(b[4:6], 'big') / 1000:.3f} V"),
                (LOCK_STATUS_CMD,   "State",         lambda b: "LOCKED" if b[4] else "UNLOCKED"),
            ]

            cell_size = parallel = 0
            for cmd, key, parser in cmds:
                resp = self._send_xgt_command(cmd)
                if resp and len(resp) > 5:
                    val = parser(resp)
                    if key == "Capacity":
                        if "Capacity" in data and isinstance(val, int) and val > 100:
                            cell_size = val
                        else:
                            parallel = val
                        if cell_size and parallel:
                            data["Capacity"] = f"{cell_size * parallel} mAh"
                    else:
                        data[key] = val

            # 10 cell voltages
            cell_cmd = bytearray([0x33, 0x23, 0x03, 0xC0, 0x00, 0x00, 0x00, 0xCC])
            voltages = []
            for i in range(1, 11):
                cell_cmd[4] = (LOOKUP[(i*2) & 0x0F] << 4) | LOOKUP[(i*2) >> 4]
                cell_cmd[1] = (LOOKUP[(i*2 + 194) & 0x0F] << 4) | LOOKUP[(i*2 + 194) >> 4]
                resp = self._send_xgt_command(bytes(cell_cmd))
                if resp and len(resp) > 5:
                    v = int.from_bytes(resp[4:6], 'big') / 1000.0
                    data[f"Cell {i} Voltage"] = f"{v:.3f} V"
                    voltages.append(v)

            if voltages:
                data["Cell Voltage Difference"] = f"{max(voltages) - min(voltages):.3f} V"

            self._insert_battery_data(data)
            self._log("[XGT] Battery data read successfully")

        except Exception as e:
            QMessageBox.critical(self, "Read Error", f"Communication failed:\n{e}")

    def _on_read_static_click(self):
        self._read_full_battery()   # model is included

    def _on_read_data_click(self):
        self._read_full_battery()

    def _reset_battery(self):
        if not self._require_interface():
            return
        try:
            self._send_xgt_command(RESET_CMD)
            time.sleep(0.01)
            self._send_xgt_command(RESET_CMD1)
            QMessageBox.information(self, "Reset", "Lockout reset commands sent!\n\nWarning: this can make the battery unusable if misused!")
        except Exception as e:
            QMessageBox.critical(self, "Reset Error", str(e))

    # ── tree helpers (exact same as LXT) ─────────────────────────────────────
    def _insert_battery_data(self, data: dict):
        self.tree.clear()
        for i, (k, v) in enumerate(data.items()):
            item = QTreeWidgetItem([k, str(v)])
            self.tree.addTopLevelItem(item)

    def _copy_to_clipboard(self):
        items = self.tree.selectedItems()
        if not items:
            items = [self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount())]
        text = "\n".join(f"{item.text(0)}: {item.text(1)}" for item in items if item)
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "Copied", "Selected data copied to clipboard.")

    def _clear_data(self):
        self._insert_battery_data(INITIAL_DATA.copy())

    def closeEvent(self, event):
        # Optional: switch back to LXT mode when closing module
        try:
            if self.interface and self.interface.serial.is_open:
                self.interface.serial.write(b'\xFF\x00')
                self.interface.serial.flush()
        except:
            pass
        event.accept()
