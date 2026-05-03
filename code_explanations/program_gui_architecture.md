# program_gui.py — Architecture Reference

---

## CLASS MAP (top-level overview)

```
program_gui.py
│
├── ErrorDialog(QDialog)              — scarlet modal, display-only
├── WarningDialog(QDialog)            — gold modal, display-only, non-blocking (open())
├── Credentials_and_mode_Dialog(QDialog)  — collects username, password, mode
├── TransferModeDialog(QDialog)       — choose Sequential vs Threaded transfer
├── SummaryDialog(QDialog)            — gold modal, displays update summary
├── PushCommandsDialog(QDialog)       — 3-phase CLI command push window
│
├── MainWindow(QMainWindow)           — root application window
│
├── ShowDevicesWorker(QObject)        — background: reads Excel, scans devices
├── PushCommandsWorker(QObject)       — background: pushes CLI commands
└── UpdateWorker(QObject)             — background: SCP transfer + install pipeline
```

---

## ErrorDialog / WarningDialog / SummaryDialog

These three are **display-only dialogs** with no internal state. They are
instantiated and called with `.exec_()` or `.open()` by their callers and
dismissed by the user via the X button.

```
Caller                      Dialog
──────                      ──────
MainWindow                ──► ErrorDialog(title, message)
  _on_show_devices_error        __init__  →  _setup_ui(title, message)
  _on_update_error                             └── QVBoxLayout
  ShowDevicesWorker.run                              ├── QHBoxLayout  [icon + title]
                                                     ├── QFrame       [scarlet separator]
                                                     ├── QLabel       [message body]
                                                     └── QLabel       [hint footer]

MainWindow                ──► WarningDialog(title, message)   [same layout as Error but gold]
  _on_update_warning            .open()  ← non-blocking, WA_DeleteOnClose set

MainWindow                ──► SummaryDialog(title, message)   [same layout but gold header]
  _on_update_finished           .exec_()
```

---

## Credentials_and_mode_Dialog

Collects username, password, and mode before any device scan starts.

```
WIDGET TREE
───────────
QDialog (470 × 360, fixed)
 └── QVBoxLayout
      ├── QLabel              "Enter credentials"  (gold header)
      ├── QFrame              scarlet separator
      ├── QLineEdit           username input
      ├── QHBoxLayout
      │    ├── QLineEdit      password input  (echo mode = Password)
      │    └── QPushButton    eye icon toggle  ──► _toggle_password_visibility()
      ├── QLabel              "Select mode"
      ├── QHBoxLayout
      │    ├── QPushButton    "UPDATER"        ──► _on_mode_selected('updater')
      │    └── QPushButton    "COMMAND PUSHER" ──► _on_mode_selected('command_pusher')
      └── QPushButton         "SUBMIT"         ──► _on_submit()  →  self.accept()

METHOD FLOW
───────────
__init__  →  _setup_ui()
               └── connects textChanged on both fields → _update_submit_state()
               └── connects mode buttons → _on_mode_selected()
               └── connects SUBMIT → _on_submit()

_toggle_password_visibility()   ← eye button clicked
    toggles QLineEdit.echoMode between Normal / Password

_on_mode_selected(mode)         ← UPDATER or COMMAND PUSHER button clicked
    stores self.selected_mode
    highlights the chosen button, dims the other
    calls _update_submit_state()

_update_submit_state()          ← called on every text change and mode selection
    enables SUBMIT only when username, password, and mode are all set

_on_submit()                    ← SUBMIT clicked
    calls self.accept()  →  exec_() returns QDialog.Accepted to caller

GETTERS (called by MainWindow after exec_()):
    get_credentials()  →  (username, password)
    get_mode()         →  'updater' or 'command_pusher'
```

---

## TransferModeDialog

Simple two-button choice shown before an update starts. Only appears when
more than one device is selected (MainWindow skips it for single-device runs).

```
WIDGET TREE
───────────
QDialog (480 × ~300, fixed)
 └── QVBoxLayout
      ├── QLabel       warning text about transfer modes
      ├── QPushButton  "SEQUENTIAL"  ──► _on_sequential()  →  accept()
      └── QPushButton  "THREADED"    ──► _on_threaded()    →  accept()

METHOD FLOW
───────────
__init__  →  _setup_ui()

_on_sequential()   sets self.selected_mode = 'sequential'  →  self.accept()
_on_threaded()     sets self.selected_mode = 'threaded'    →  self.accept()

CALLER:
    MainWindow._on_start()
        if len(selected_indices) == 1:  transfer_mode = 'sequential'  (dialog skipped)
        else:  mode_dialog.exec_()  →  transfer_mode = mode_dialog.selected_mode
```

---

## PushCommandsDialog

Three-phase dialog managed by a QStackedWidget. Phases are pages; transitions
are `_stack.setCurrentIndex(n)`.

```
WIDGET TREE
───────────
QDialog (700 × 520, fixed)
 └── QVBoxLayout  (zero margins — stack fills edge-to-edge)
      └── QStackedWidget  (_stack)
           │
           ├── Page 0 — PHASE INPUT  (_build_phase_input)
           │    └── QVBoxLayout
           │         ├── QLabel          "Enter commands:"  (gold, #cmdLabel)
           │         ├── QPlainTextEdit  _editor — no wrap, h+v scroll as needed
           │         └── QHBoxLayout
           │              └── QPushButton  "PUSH"  (#pushBtn, green, starts disabled)
           │
           ├── Page 1 — PHASE PUSHING  (_build_phase_pushing)
           │    └── QVBoxLayout
           │         ├── stretch
           │         ├── QLabel  "Pushing..."  (centered, #statusLabel)
           │         └── stretch
           │
           └── Page 2 — PHASE OUTPUT  (_build_phase_output)
                └── QVBoxLayout
                     ├── QLabel       "COMMAND OUTPUT"  (gold, #outputHeader)
                     ├── QFrame       scarlet separator
                     ├── QScrollArea  _tab_scroll  (h-scroll only, 46px fixed height)
                     │    └── QWidget  _tab_inner
                     │         └── QHBoxLayout  _tab_layout
                     │              ├── QPushButton  "SW2"  (#tabBtn / #tabBtnActive)
                     │              ├── QPushButton  "SW3"
                     │              ├── ...one per device...
                     │              └── stretch  (keeps buttons left-aligned)
                     ├── QPlainTextEdit  _output_area  (read-only, both scrollbars)
                     └── QHBoxLayout
                          └── QPushButton  "PUSH MORE"  (#pushMoreBtn, green)

METHOD FLOW
───────────
SETUP
  __init__
      └── _setup_ui()
               ├── _build_phase_input()    → Page 0 widget
               ├── _build_phase_pushing()  → Page 1 widget
               └── _build_phase_output()   → Page 2 widget

  set_push_context(selected_df, username, password)     ← called by MainWindow BEFORE exec_()
      stores _selected_df, _username, _password
      └── _build_device_tabs()   builds tab buttons into _tab_layout

PHASE 0 → 1 (user clicks PUSH)
  _editor.textChanged ──► _on_text_changed()
      enforces MAX_LINES cap (blockSignals pattern)
      enables/disables _btn_push based on content

  _btn_push.clicked ──► _on_push_clicked()
      sets _pushing = True  (blocks closeEvent)
      _stack.setCurrentIndex(_PHASE_PUSHING)
      creates PushCommandsWorker + QThread
      wires: thread.started → worker.run
             worker.finished → _on_push_finished(results)
             worker.error    → _on_push_error(error_msg)
             worker.finished → thread.quit → worker.deleteLater + thread.deleteLater
      thread.start()

PHASE 1 → 2 (worker finishes)
  worker.finished ──► _on_push_finished(results)
      _pushing = False  (unblocks closeEvent)
      stores _results = {df_index: result_str}
      _build_device_tabs()  (rebuilds tabs for the new push cycle)
      _select_device_tab(first_index)
      _stack.setCurrentIndex(_PHASE_OUTPUT)

  worker.error ──► _on_push_error(error_msg)
      _pushing = False
      fills _results with WORKER ERROR string for every device
      same tab build + auto-select + page switch as finished

PHASE 2 — tab interaction
  tab button clicked ──► _select_device_tab(index)
      updates objectName on all buttons (tabBtn / tabBtnActive)
      calls unpolish + polish to force stylesheet re-evaluation
      └── _show_device_output(index)
               reads _results[index]
               reads hostname + IP from _selected_df via .at[]
               sets _output_area.setPlainText(formatted string)

PHASE 2 → 0 (user clicks PUSH MORE)
  _btn_push_more.clicked ──► _on_push_more_clicked()
      _editor.clear()
      _btn_push.setEnabled(False)
      _stack.setCurrentIndex(_PHASE_INPUT)

CLOSE GUARD
  closeEvent(event)
      _pushing == True  →  event.ignore()   (phase 1 only)
      _pushing == False →  event.accept()   (phases 0 and 2)

CALLER (MainWindow):
  _on_push_commands()
      builds selected_df from checked table rows
      dlg = PushCommandsDialog(self)
      dlg.set_push_context(selected_df, username, password)
      dlg.exec_()   ← blocking; worker runs inside dialog on its own thread
```

---

## PushCommandsWorker

Thin background worker. All device-level logic is in `device_cli_ops`.

```
Signals:  finished(list)  — list[(df_index, result_str)]
          error(str)       — unhandled exception string

__init__(selected_df, username, password, commands)
    stores args

run()   ← called by QThread.started signal
    results = device_cli_ops.push_commands_all(...)
    self.finished.emit(results)
    — OR —
    self.error.emit(str(e))   on exception

WIRED BY: PushCommandsDialog._on_push_clicked()
```

---

## ShowDevicesWorker

Heavy background worker. Orchestrates the full device scan pipeline.
Each stage calls an `excel_and_data_ops` function, then calls `_update_tracker`
on error.

```
Signals:  finished(object, object)  — (eligible_df, valid_devices_df)
          error(str)
          progress(str)

__init__(excel_file, sheet_name, username, password, mode)

run()
  ├── excel_and_data_ops.load_valid_devices_df()
  ├── _update_tracker(valid_devices_df)  ←─────────────────────────────┐
  ├── populate_status_and_auth_status_column()                          │
  ├── populate_restconf_status_column()     (mode='updater' only)       │  called after
  ├── populate_scp_status_column()          (mode='updater' only)       │  every populate
  ├── populate_current_version_column()     (mode='updater' only)       │  stage + final
  ├── populate_flash_free_space_column()    (mode='updater' only)       │  Excel write
  ├── populate_ios_image_path_column()      (mode='updater' only)       │
  ├── populate_needs_update_column()        (mode='updater' only)       │
  ├── _update_tracker(valid_devices_df)  ←─────────────────────────────┘
  └── get_eligible_devices_df()
       → self.finished.emit(eligible_df, valid_devices_df)

_update_tracker(valid_devices_df)
    calls excel_and_data_ops.update_excel_tracker()
    on any error → self.error.emit(message)  and returns False
    returns True on success
    caller does:  if not self._update_tracker(...): return

WIRED BY: MainWindow._on_show_devices()
    worker.finished → _on_show_devices_done(eligible_df, valid_devices_df)
    worker.error    → _on_show_devices_error(message)
    worker.progress → _update_loading_message(message)
```

---

## UpdateWorker

Orchestrates the multi-stage firmware update. Calls `_update_tracker` after
every stage to persist progress. Non-fatal tracker failures emit `warning`.

```
Signals:  finished(str)   — final summary string
          error(str)       — fatal failure string
          progress(str)    — stage description
          warning(str)     — non-fatal Excel write failure

__init__(excel_file, sheet_name, valid_devices_df, selected_indices,
         username, password, transfer_mode)

run()
  ├── PART 3: device_file_transfer_ops.scp_transfer_all()
  │    └── _update_tracker()  ──► warning.emit() on failure (non-fatal)
  ├── PART 4: device_cli_ops.install_ios_image_all()
  │    └── _update_tracker()
  ├── PART 5: post-reload poll + device_cli_ops.commit_ios_install_all()
  │    └── _update_tracker()
  ├── PART 6: device_cli_ops.remove_inactive_ios_all()
  │    └── _update_tracker()
  └── _generate_summary()  →  self.finished.emit(summary)

_update_tracker()
    calls excel_and_data_ops.update_excel_tracker()
    on error → self.warning.emit(message)   ← does NOT halt the update
    (contrast with ShowDevicesWorker._update_tracker which halts on error)

_generate_summary()
    builds a per-device result table from valid_devices_df
    returns formatted string → passed to finished.emit()

WIRED BY: MainWindow._on_start()
    worker.finished → _on_update_finished(summary)  →  SummaryDialog
    worker.error    → _on_update_error(msg)          →  ErrorDialog + UI restore
    worker.warning  → _on_update_warning(msg)        →  WarningDialog.open()  (non-blocking)
    worker.progress → _on_update_progress(msg)       →  _update_loading_message()
```

---

## MainWindow

Root window. Owns all UI state, all worker lifecycles, and all signal wiring.

```
WIDGET TREE (static layout)
───────────────────────────
QMainWindow (860 × 610, fixed)
 └── QWidget  centralWidget
      └── QVBoxLayout  root
           ├── QLabel            app title
           ├── QLabel            subtitle
           ├── QFrame            shimmer separator
           ├── QHBoxLayout       top row
           │    ├── QLineEdit    sheet_input  (Excel sheet name)
           │    └── QPushButton  btn_show     "SHOW DEVICES"
           ├── QHBoxLayout       table header row
           │    ├── QLabel       "ELIGIBLE DEVICES"
           │    └── QPushButton  btn_select_all  (positioned dynamically)
           ├── QTableWidget      table  (device rows with checkboxes)
           └── QHBoxLayout       _bottom_layout  (dynamic — rebuilt per mode)
                ← UPDATER mode:        [CANCEL UPDATE]  [START UPDATE]
                ← COMMAND PUSHER mode: [PUSH COMMANDS]

DYNAMIC BOTTOM BUTTONS
───────────────────────
_build_bottom_buttons(mode)
    'updater'        → btn_cancel + btn_start
    'command_pusher' → btn_push_commands
    connects clicked signals

_clear_bottom_buttons()
    calls widget.disconnect() THEN widget.deleteLater() on every item
    (disconnect before deleteLater prevents stale signal firing during deferred deletion)

KEY METHOD INTERACTIONS
────────────────────────
btn_show.clicked
 └── _on_show_devices()
      ├── Credentials_and_mode_Dialog.exec_()  → stores credentials + mode
      ├── checks Excel file not locked
      ├── _clear_bottom_buttons()
      ├── _show_loading_in_table()
      └── ShowDevicesWorker + QThread  (wires finished/error/progress)

ShowDevicesWorker.finished
 └── _on_show_devices_done(eligible_df, valid_devices_df)
      ├── stores _eligible_df, _valid_devices_df
      ├── _hide_loading_in_table()
      ├── _populate_table(eligible_df)
      └── _build_bottom_buttons(mode)

ShowDevicesWorker.error
 └── _on_show_devices_error(message)
      ├── _hide_loading_in_table()
      └── ErrorDialog.exec_()

btn_start.clicked  (UPDATER mode)
 └── _on_start()
      ├── collects selected_indices from table checkboxes
      ├── 1 device → transfer_mode = 'sequential'
        2+ devices → TransferModeDialog.exec_()
      ├── locks UI
      ├── _show_loading_in_table()
      └── UpdateWorker + QThread  (wires finished/error/warning/progress)

btn_push_commands.clicked  (COMMAND PUSHER mode)
 └── _on_push_commands()
      ├── collects selected_indices from table checkboxes
      ├── builds selected_df from _valid_devices_df
      └── PushCommandsDialog(self).exec_()
           (worker lifecycle lives entirely inside the dialog)

TABLE HELPERS
─────────────
_populate_table(eligible_df)           fills rows with device data + checkboxes
_get_checkbox(row)                     returns QCheckBox widget for a given row
_update_button_states()                enables/disables btn_start / btn_push_commands
                                       based on checkbox selection count
_update_show_button_state()            enables/disables btn_show based on sheet_input content
_toggle_select_all()                   checks/unchecks all visible checkboxes
_on_header_clicked(col)                sorts table by column
_distribute_columns_evenly()           spreads column widths equally
_fit_columns_for_updater()             applies fixed widths for updater mode
_reposition_select_all()               moves btn_select_all over the checkbox column header

LOADING OVERLAY
───────────────
_show_loading_in_table(message)        replaces table content with a single spanning "Loading..." row
_update_loading_message(message)       updates that row's text (called by worker.progress signals)
_hide_loading_in_table()               restores the table to its normal state

UI RESTORE
──────────
_restore_ui_after_update()             re-enables all inputs and buttons after UpdateWorker finishes

CANCEL
──────
btn_cancel.clicked
 └── _on_cancel_clicked()
      sets device_cli_ops.cancel_event
      calls cancel_active_transfers_all() directly (not via worker)
```

---

## Signal Flow Summary

```
USER ACTION                     MAIN THREAD                      BACKGROUND THREAD
───────────                     ───────────                      ─────────────────
Click SHOW DEVICES          →   ShowDevicesWorker.start()   →   run() scans devices
                            ←   progress(str)               ←   self.progress.emit()
                            ←   finished(df, df)            ←   self.finished.emit()
                            ←   error(str)                  ←   self.error.emit()

Click START UPDATE          →   UpdateWorker.start()        →   run() SCP + install
                            ←   progress(str)               ←   self.progress.emit()
                            ←   warning(str)                ←   self.warning.emit()   [Excel fail]
                            ←   finished(str)               ←   self.finished.emit()
                            ←   error(str)                  ←   self.error.emit()

Click PUSH in dialog        →   PushCommandsWorker.start()  →   run() push commands
                            ←   finished(list)              ←   self.finished.emit(results)
                            ←   error(str)                  ←   self.error.emit()
```
