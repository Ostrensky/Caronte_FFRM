# --- FILE: app/video_splash.py ---

import os
from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel
from PySide6.QtMultimedia import QMediaPlayer, QAudioOutput
from PySide6.QtMultimediaWidgets import QVideoWidget
from PySide6.QtCore import QUrl, Signal, Qt, QTimer

class VideoSplashScreen(QWidget):
    finished = Signal()

    def __init__(self, video_path, width=800, height=600):
        super().__init__()
        
        # 1. VISUAL SETUP
        # Removing "SplashScreen" flag sometimes fixes rendering issues on Windows
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.setAttribute(Qt.WidgetAttribute.WA_DontCreateNativeAncestors, False)
        
        self.resize(width, height)
        self.center_on_screen()
        self.setStyleSheet("background-color: black;") # Force black bg

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 2. MEDIA PLAYER SETUP
        self.video_widget = QVideoWidget()
        layout.addWidget(self.video_widget)

        self.player = QMediaPlayer()
        self.audio_output = QAudioOutput()
        
        self.player.setAudioOutput(self.audio_output)
        self.player.setVideoOutput(self.video_widget)
        self.audio_output.setVolume(0.5) 

        # 3. DEBUG CONNECTIONS
        self.player.mediaStatusChanged.connect(self.handle_media_status)
        self.player.playbackStateChanged.connect(self.handle_state_change)
        self.player.errorOccurred.connect(self.handle_error)

        # 4. LOAD SOURCE
        if not os.path.exists(video_path):
            print(f"[SPLASH] ❌ File not found: {video_path}")
            QTimer.singleShot(100, self.finished.emit)
            return

        abs_path = os.path.abspath(video_path)
        print(f"[SPLASH] Loading video: {abs_path}")
        self.player.setSource(QUrl.fromLocalFile(abs_path))

    def start(self):
        self.show()
        # Small delay to ensure window is mapped before playing
        QTimer.singleShot(100, self.player.play)

    def handle_media_status(self, status):
        print(f"[SPLASH] Media Status: {status}")
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            print("[SPLASH] Video finished.")
            self.stop_and_finish()
        elif status == QMediaPlayer.MediaStatus.InvalidMedia:
            print("[SPLASH] ❌ Invalid Media (Codec issue?)")
            self.stop_and_finish()

    def handle_state_change(self, state):
        print(f"[SPLASH] Playback State: {state}")

    def handle_error(self):
        err_str = self.player.errorString()
        print(f"[SPLASH] ❌ PLAYER ERROR: {err_str}")
        self.stop_and_finish()

    def stop_and_finish(self):
        print("[SPLASH] Closing splash.")
        self.player.stop()
        self.close()
        self.finished.emit()

    def mousePressEvent(self, event):
        print("[SPLASH] User skipped.")
        self.stop_and_finish()

    def center_on_screen(self):
        from PySide6.QtGui import QGuiApplication
        screen = QGuiApplication.primaryScreen().availableGeometry()
        self.move((screen.width() - self.width()) // 2, (screen.height() - self.height()) // 2)