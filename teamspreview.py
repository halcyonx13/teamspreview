import pyglet
from pyglet.gl import GL_RGBA
from pyglet.window import mouse, key
from windows_capture import WindowsCapture, Frame, InternalCaptureControl
import time
import json
from dataclasses import dataclass, asdict, fields
from pathlib import Path
import win32gui
import ctypes

VER = 0.2

CONFIG_PATH = Path.home() / ".teams_preview_config.json"
SW_HIDE = 0
SW_SHOW = 5
HWND_TOPMOST = -1
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001


class WindowChooser:
    def __init__(self, window, font_name="Consolas"):
        self.window = window
        self.font_name = font_name

        self.items: list[str] = []
        self.labels: list[pyglet.text.Label] = []
        self.info_label: pyglet.text.Label | None = None
        self.menu_keys_label: pyglet.text.Label | None = None
        self.mouse_control_label: pyglet.text.Label | None = None

        self.current_index: int = 0
        self.active: bool = False

        self._last_window_size: tuple[int, int] | None = None

    def show(self, items: list[str], initial_title: str | None = None):
        self.items = items
        self.active = True

        self.current_index = 0
        if initial_title:
            for i, title in enumerate(items):
                if title == initial_title:
                    self.current_index = i
                    break

        self._rebuild_labels()

    def hide(self):
        """Deactivate the chooser and drop references to labels."""
        self.active = False
        self.items = []
        self.labels = []
        self.info_label = None
        self.menu_keys_label = None
        self.mouse_control_label = None

    def get_selected_title(self) -> str | None:
        if not self.items:
            return None
        return self.items[self.current_index]

    def draw(self):
        if not self.active:
            return

        size = (self.window.width, self.window.height)
        if size != self._last_window_size:
            self._rebuild_labels()

        self.window.clear()

        if self.info_label:
            self.info_label.draw()
        if self.menu_keys_label:
            self.menu_keys_label.draw()
        if self.mouse_control_label:
            self.mouse_control_label.draw()

        for i, label in enumerate(self.labels):
            if i == self.current_index:
                label.color = (0, 255, 255, 255)
            else:
                label.color = (255, 255, 255, 255)
            label.draw()

    def handle_key_press(self, symbol, modifiers) -> bool:
        if not self.active:
            return False

        if symbol == key.DOWN and self.items:
            self.current_index = (self.current_index + 1) % len(self.items)
            return True

        if symbol == key.UP and self.items:
            self.current_index = (self.current_index - 1) % len(self.items)
            return True

        return False

    def _rebuild_labels(self):
        self._last_window_size = (self.window.width, self.window.height)

        self.labels = []

        spacing = 30
        y_offset = 40

        self.info_label = pyglet.text.Label(
            "Choose the Teams meeting window:",
            font_name=self.font_name, font_size=24,
            x=25, y=self.window.height - 25 - y_offset,
            anchor_x="left", anchor_y="top",
            color=(255, 255, 255, 255))

        list_top = self.window.height - 70 - y_offset
        x = 60
        y = list_top

        for title in self.items:
            label = pyglet.text.Label(
                title,
                font_name=self.font_name, font_size=18,
                x=x, y=y,
                anchor_x="left", anchor_y="top",
                color=(255, 255, 255, 255))
            self.labels.append(label)
            y -= spacing

        self.menu_keys_label = pyglet.text.Label(
            "[Space] show this menu   [F10] reset pan/zoom   [F11] show FPS   [F12] colour mode",
            font_name=self.font_name, font_size=14,
            x=25, y=40,
            anchor_x="left", anchor_y="bottom",
            color=(200, 200, 200, 255))

        self.mouse_control_label = pyglet.text.Label(
            "[LMB]   move window   [RMB] pan   [WHEEL] zoom",
            font_name=self.font_name, font_size=14,
            x=25, y=20,
            anchor_x="left", anchor_y="bottom",
            color=(200, 200, 200, 255))


@dataclass
class AppConfig:
    last_window_title: str = ""

    pan_x: float = 0.0
    pan_y: float = 0.0
    zoom: float = 1.0

    show_fps: bool = False
    color_toggle: bool = False

    window_height: int = 720
    window_width: int = 1280
    window_pos_x: int = 100
    window_pos_y: int = 100

    fps: int = 30
    titlefilter: str = "Teams"
    vsync: bool = False

    @classmethod
    def load(cls) -> "AppConfig":
        try:
            raw = CONFIG_PATH.read_text(encoding="utf-8")
            data = json.loads(raw)
        except FileNotFoundError:
            # no config yet, use defaults
            return cls()
        except (OSError, json.JSONDecodeError):
            # can't read or parse, fall back to defaults
            return cls()

        # only keep known fields, ignore extras
        field_names = {f.name for f in fields(cls)}
        kwargs = {k: v for k, v in data.items() if k in field_names}

        return cls(**kwargs)

    def save(self) -> None:
        try:
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        except OSError:
            # if can't create directory, silently ignore
            return

        data = asdict(self)
        tmp_path = CONFIG_PATH.with_suffix(".tmp")

        # write to temp file then rename
        tmp_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        tmp_path.replace(CONFIG_PATH)

_latest_frame = None
_latest_size = (0, 0)
_last_frame_time = None
_window_visible = None
control = None

capture = None
capture_hwnd = None

config = AppConfig.load()
last_window_title = config.last_window_title or None


def get_windows_callback(_hwnd, strings):
    if not win32gui.IsWindowVisible(_hwnd):
        return True

    window_title = win32gui.GetWindowText(_hwnd)
    left, top, right, bottom = win32gui.GetWindowRect(_hwnd)
    width = right - left
    height = bottom - top

    if window_title and width > 0 and height > 0:
        strings.append(window_title)

    return True


def get_windows(titlefilter=None):
    win_list = []
    win32gui.EnumWindows(get_windows_callback, win_list)

    windows_list = []
    for title in win_list:
        if title == "Teams Preview":
            continue
        if titlefilter is None or titlefilter in title:
            windows_list.append(title)

    return windows_list


def open_window_menu():
    titles = get_windows(config.titlefilter)

    if not titles:
        print("No Teams windows found")
        return

    chooser.show(titles, initial_title=config.last_window_title or None)


FRAME_TIMEOUT_SECONDS = 1.0

current_windows = get_windows(config.titlefilter)
window_found = False
if last_window_title is not None:
    for w in current_windows:
        if w == last_window_title:
            window_found = True

if not window_found:
    window_w = 960
    window_h = 540
    window = pyglet.window.Window(resizable=True, width=window_w, height=window_h, caption="Teams Preview", style=pyglet.window.Window.WINDOW_STYLE_BORDERLESS, vsync=config.vsync)
    chooser = WindowChooser(window)
    open_window_menu()
    center_x = (window.screen.width // 2) - (window_w // 2)
    center_y = window.screen.height // 2 - (window_h // 2)
    window.set_location(center_x, center_y)
else:
    window = pyglet.window.Window(resizable=True, width=config.window_width, height=config.window_height, caption="Teams Preview", style=pyglet.window.Window.WINDOW_STYLE_BORDERLESS, vsync=config.vsync)
    chooser = WindowChooser(window)
    window.set_location(config.window_pos_x, config.window_pos_y)

fps_display = pyglet.window.FPSDisplay(window=window)
self_hwnd = window._hwnd
win32gui.SetWindowPos(self_hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)  # stay on top

no_frames_label = pyglet.text.Label("Unminimise the Teams meeting window", font_name="Consolas", font_size=18, x=window.width // 2, y=window.height // 2, anchor_x="center", anchor_y="center", color=(255, 255, 255, 255))

_window_texture = None
_frame_image = None
_frame_image_format = None
_frame_buffer_type = None

pan_x = config.pan_x
pan_y = config.pan_y
zoom = config.zoom

colour_mode_toggle = config.color_toggle
fps_toggle = config.show_fps

min_zoom = 0.2
max_zoom = 5.0

@window.event
def on_resize(width, height):

    if not chooser.active:
        config.window_width = int(width)
        config.window_height = int(height)

    no_frames_label.x = width // 2
    no_frames_label.y = height // 2
    return pyglet.window.Window.on_resize(window, width, height)


@window.event
def on_draw():
    window.clear()

    if chooser.active:
        chooser.draw()
        return
    else:
        has_frame, frame, (frame_w, frame_h) = have_recent_frame()
        if not has_frame:
            no_frames_label.x = window.width // 2
            no_frames_label.y = window.height // 2
            no_frames_label.draw()
            return

        if window.width == 0 or window.height == 0 or frame_w == 0 or frame_h == 0: return

        global _window_texture, _frame_image, _frame_image_format, _frame_buffer_type

        scale = zoom * (window.height / float(frame_h))
        display_w = frame_w * scale
        display_h = frame_h * scale

        base_x = (window.width - display_w) / 2.0
        base_y = 0

        x = base_x + pan_x
        y = base_y + pan_y

        if _window_texture is None or _window_texture.width != frame_w or _window_texture.height != frame_h:
            _window_texture = pyglet.image.Texture.create(frame_w, frame_h, internalformat=GL_RGBA)

            # flip texcoords vertically once
            tc = list(_window_texture.tex_coords)
            for i in range(1, len(tc), 3):
                tc[i] = 1.0 - tc[i]
            _window_texture.tex_coords = tuple(tc)

            _frame_image = None
            _frame_image_format = None
            _frame_buffer_type = None

        color_mode = 'RGBA' if colour_mode_toggle else 'BGRA'
        nbytes = frame_w * frame_h * 4

        if _frame_buffer_type is None:
            _frame_buffer_type = ctypes.c_ubyte * nbytes

        buf = _frame_buffer_type.from_address(frame.ctypes.data)

        if _frame_image is None or _frame_image.width != frame_w or _frame_image.height != frame_h or _frame_image_format != color_mode:
            _frame_image = pyglet.image.ImageData(frame_w, frame_h, color_mode, buf, pitch=frame_w * 4)
            _frame_image_format = color_mode
        else:
            _frame_image.set_data(color_mode, frame_w * 4, buf)

        _window_texture.blit_into(_frame_image, 0, 0, 0)
        _window_texture.blit(x, y, width=display_w, height=display_h)

        if fps_toggle:
            fps_display.draw()


@window.event
def on_move(x, y):
    if not chooser.active:
        config.window_pos_x = int(x)
        config.window_pos_y = int(y)
    else:
        config.window_pos_x = int(x)
        config.window_pos_y = int(y)


@window.event
def on_mouse_scroll(x, y, scroll_x, scroll_y):
    global zoom, pan_x, pan_y

    if _latest_frame is None: return
    frame_w, frame_h = _latest_size

    old_zoom = zoom

    zoom_factor = 1.1
    if scroll_y > 0:
        zoom *= zoom_factor
    elif scroll_y < 0:
        zoom /= zoom_factor

    zoom = max(min_zoom, min(max_zoom, zoom))

    if zoom == old_zoom: return

    def compute_transform(z):
        base_scale = window.height / float(frame_h)
        s = z * base_scale
        display_w = frame_w * s
        display_h = window.height
        bx = (window.width - display_w) / 2.0
        by = 0.0
        return s, bx, by

    old_scale, old_base_x, old_base_y = compute_transform(old_zoom)
    new_scale, new_base_x, new_base_y = compute_transform(zoom)

    if old_scale == 0: return

    factor = new_scale / old_scale

    pan_x = x - new_base_x - factor * (x - old_base_x - pan_x)
    pan_y = y - new_base_y - factor * (y - old_base_y - pan_y)

    config.pan_x = pan_x
    config.pan_y = pan_y
    config.zoom = zoom


@window.event
def on_mouse_press(x, y, button, modifiers):

    # deal with moving the window
    if button == mouse.LEFT:
        win32gui.ReleaseCapture()
        win32gui.SendMessage(self_hwnd, 0xA1, 2, 0)   # WM_NCLBUTTONDOWN + HTCAPTION


@window.event
def on_mouse_drag(x, y, dx, dy, buttons, modifiers):
    global pan_x, pan_y
    if buttons & mouse.RIGHT:
        pan_x += dx
        pan_y += dy
        config.pan_x = pan_x
        config.pan_y = pan_y


@window.event
def on_key_press(symbol, modifiers):
    global colour_mode_toggle, fps_toggle, capture, capture_hwnd, control, pan_x, pan_y, zoom

    if chooser.active:
        handled = chooser.handle_key_press(symbol, modifiers)
        if handled:
            return

        if symbol == key.ENTER:
            selected_title = chooser.get_selected_title()
            if selected_title:
                start_capture(selected_title)
                config.last_window_title = selected_title
            chooser.hide()
            window.set_location(config.window_pos_x, config.window_pos_y)
            window.set_size(config.window_width, config.window_height)
            return

        if symbol == key.ESCAPE:
            config.last_window_title = ''
            config.save()
            pyglet.app.exit()
            return

    if symbol == pyglet.window.key.F12:
        colour_mode_toggle = not colour_mode_toggle
        config.color_toggle = bool(colour_mode_toggle)

    if symbol == pyglet.window.key.F11:
        fps_toggle = not fps_toggle
        config.show_fps = bool(fps_toggle)

    if symbol == pyglet.window.key.SPACE:
        if control is not None:
            control.stop()
            control = None

        capture = None
        capture_hwnd = None
        open_window_menu()
        return

    if symbol == key.F10:
        zoom = 1.0
        pan_x, pan_y = 0, 0
        config.pan_x = pan_x
        config.pan_y = pan_y
        config.zoom = zoom


def have_recent_frame():
    frame = _latest_frame
    width, height = _latest_size
    last_time = _last_frame_time

    if frame is None or last_time is None:
        return False, None, (width, height)

    # if the window is minimised
    if time.monotonic() - last_time > FRAME_TIMEOUT_SECONDS:
        return False, frame, (width, height)

    return True, frame, (width, height)


def check_active_window(dt):
    global _window_visible
    active_hwnd = win32gui.GetForegroundWindow()
    should_be_visible = active_hwnd != capture_hwnd
    if should_be_visible != _window_visible:
        win32gui.ShowWindow(self_hwnd, SW_SHOW if should_be_visible else SW_HIDE)
        _window_visible = should_be_visible


def on_closed():
    print("Capture Session Closed")
    pyglet.app.exit()


def on_frame_arrived(frame: Frame, capture_control: InternalCaptureControl):
    global _latest_frame, _latest_size, _last_frame_time

    fb = frame.frame_buffer
    if fb is None:
        return

    fh, fw, _ = fb.shape

    _latest_frame = fb.copy()
    _latest_size = (fw, fh)
    _last_frame_time = time.monotonic()


def start_capture(win_title):
    global control, capture_hwnd, capture
    capture = WindowsCapture(cursor_capture=False, draw_border=None, monitor_index=None, window_name=win_title)
    capture.frame_handler = on_frame_arrived
    capture.closed_handler = on_closed
    capture_hwnd = win32gui.FindWindow(None, win_title)
    control = capture.start_free_threaded()


if window_found:
    start_capture(last_window_title)
    chooser.hide()
    config.last_window_title = last_window_title
    window.set_size(config.window_width, config.window_height)
    window.set_location(config.window_pos_x, config.window_pos_y)

pyglet.clock.schedule_interval(check_active_window, 1/10)

pyglet.app.run(1/config.fps)
config.save()

print("Exiting...")
