
import sys
import os
import win32com.client
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QPushButton, QProgressBar,
                             QComboBox, QFrame, QFileDialog, QMessageBox,
                             QTextEdit, QShortcut,
                             QDesktopWidget,
                             QSlider)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon, QKeySequence
from PyQt5.QtCore import QDateTime
from datetime import datetime
from PyQt5.QtWidgets import QSizeGrip
# from PyQt5.QtCore import QWIDGETSIZE_MAX


class TaskDetail:
    def __init__(self, task_id, description, required_actions, file_name):
        self.task_id = task_id
        self.description = description
        self.required_actions = required_actions
        self.file_name = file_name


class TestWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.is_always_on_top = True
        self.word_app = None
        self.source_doc = None
        self.current_task = 1
        self.total_tasks = 10
        self.task_states = {i: 'incomplete' for i in range(1, 11)}
        self.task_details = {}  # Will store TaskDetail objects
        self.save_folder = self.create_save_folder()  # Initialize save folder

        # Get screen dimensions
        self.screen_size = self.get_screen_size()
        
        
        # Set minimum window size
        self.min_width = int(self.screen_size.width() * 0.5)   # 50% of screen width
        self.min_height = int(self.screen_size.height() * 0.15)  # 15% of screen height

            # Calculate window dimensions and position
        self.window_width = int(self.screen_size.width() * 0.98)  # 98% of screen width
        self.window_height = int(self.screen_size.height() * 0.2)  # 20% of screen height
        
        # Calculate position (centered horizontally, bottom of screen)
        self.window_x = int((self.screen_size.width() - self.window_width) / 2)
        self.window_y = int(self.screen_size.height() - self.window_height)
        self.load_task_details()  # Load tasks from Excel
        self.initUI()
        
        # Position window at bottom
        self.position_window_bottom()
        

    def set_window_flags(self):
        """Set window flags based on always-on-top state"""
        if self.is_always_on_top:
            self.setWindowFlags(
                Qt.Window |
                Qt.WindowStaysOnTopHint |
                Qt.WindowTitleHint |
                Qt.CustomizeWindowHint |
                Qt.WindowCloseButtonHint |
                Qt.WindowMinimizeButtonHint
            )
        else:
            self.setWindowFlags(
                Qt.Window |
                Qt.WindowTitleHint |
                Qt.CustomizeWindowHint |
                Qt.WindowCloseButtonHint |
                Qt.WindowMinimizeButtonHint
            )
        
        # Show window after changing flags
        self.show()
        
        # Reactivate Word window if it exists
        if self.word_app:
            try:
                self.word_app.Activate()
            except:
                pass

    def toggle_always_on_top(self):
        """Toggle always-on-top state"""
        self.is_always_on_top = not self.is_always_on_top
        self.update_pin_button_style()
        self.set_window_flags()
    def position_window_bottom(self):
        """Position window at the bottom of the screen"""
        self.move(self.window_x, self.window_y)

    def update_pin_button_style(self):
        """Update pin button appearance based on state"""
        if self.is_always_on_top:
            self.pin_button.setStyleSheet('''
                QPushButton {
                    background-color: #FFD700;
                    border: none;
                    border-radius: 4px;
                    padding: 4px;
                    qproperty-text: "📌";
                }
                QPushButton:hover {
                    background-color: #FFF0AA;
                }
            ''')
            self.pin_button.setToolTip("Click to unpin window (Currently: Always on top)")
        else:
            self.pin_button.setStyleSheet('''
                QPushButton {
                    background-color: transparent;
                    border: 1px solid white;
                    border-radius: 4px;
                    padding: 4px;
                    color: white;
                    qproperty-text: "📌";
                }
                QPushButton:hover {
                    background-color: rgba(255, 255, 255, 0.1);
                }
            ''')
            self.pin_button.setToolTip("Click to pin window on top")

    def launch_word(self):
        """Launch Microsoft Word application with window positioning"""
        try:
            # Create Word application instance
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = True
            
            # Position Word window
            if hasattr(self.word_app, 'WindowState'):
                self.word_app.WindowState = 1  # Maximize Word window
            
            # Set initial file
            self.open_source_document(1)

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error launching Word: {str(e)}")

    def resizeEvent(self, event):
        """Handle window resize events"""
        super().resizeEvent(event)
        # Ensure Word window is visible if it exists
        if self.word_app:
            try:
                self.word_app.Activate()
            except:
                pass
    def create_zoom_control(self):
        """Create zoom control slider"""
        zoom_container = QHBoxLayout()
        zoom_container.setContentsMargins(0, 0, 0, 0)
        
        # Create zoom out label
        zoom_out_label = QLabel("A-")
        zoom_out_label.setStyleSheet('''
            QLabel {
                color: white;
                font-weight: bold;
                margin-right: 5px;
            }
        ''')
        
        # Create zoom slider
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setMinimum(8)    # Minimum font size
        self.zoom_slider.setMaximum(24)   # Maximum font size
        self.zoom_slider.setValue(13)      # Default font size
        self.zoom_slider.setFixedWidth(100)
        self.zoom_slider.setStyleSheet('''
            QSlider::groove:horizontal {
                border: 1px solid white;
                height: 4px;
                background: white;
                margin: 0px;
                border-radius: 2px;
            }
            QSlider::handle:horizontal {
                background: #FFD700;
                border: 1px solid #FFD700;
                width: 12px;
                height: 12px;
                margin: -4px 0;
                border-radius: 6px;
            }
            QSlider::handle:horizontal:hover {
                background: #FFF0AA;
                border: 1px solid #FFF0AA;
            }
        ''')
        self.zoom_slider.valueChanged.connect(self.update_font_size)
        
        # Create zoom in label
        zoom_in_label = QLabel("A+")
        zoom_in_label.setStyleSheet('''
            QLabel {
                color: white;
                font-weight: bold;
                margin-left: 5px;
            }
        ''')
        
        # Add widgets to zoom container
        zoom_container.addWidget(zoom_out_label)
        zoom_container.addWidget(self.zoom_slider)
        zoom_container.addWidget(zoom_in_label)
        zoom_container.addStretch()
        
        return zoom_container

    def update_font_size(self):
        """Update the font size of the description text"""
        font_size = self.zoom_slider.value()
        self.description_text.setStyleSheet(f'''
            QTextEdit {{
                background-color: transparent;
                color: white;
                border: none;
                font-size: {font_size}px;
                line-height: 1.6;
                selection-background-color: rgba(255, 215, 0, 0.3);
                selection-color: white;
            }}
            QTextEdit:focus {{
                border: none;
                outline: none;
            }}
            QScrollBar:vertical {{
                border: none;
                background: rgba(255, 255, 255, 0.1);
                width: 10px;
                margin: 0;
            }}
            QScrollBar::handle:vertical {{
                background: rgba(255, 255, 255, 0.3);
                min-height: 20px;
                border-radius: 5px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: rgba(255, 255, 255, 0.4);
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
                background: none;
            }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                background: none;
            }}
        ''')


    def load_task_details(self):
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            excel_path = os.path.join(
                current_dir, "Project1_Requirements.xlsx")

            if not os.path.exists(excel_path):
                QMessageBox.warning(
                    self, "Error", f"Task file not found: {excel_path}")
                return

            df = pd.read_excel(excel_path)

            for _, row in df.iterrows():
                task_id = row['TaskID']
                task_detail = TaskDetail(
                    task_id=task_id,
                    description=row['Description'],
                    required_actions=row['RequiredActions'].split(
                        ';') if isinstance(row['RequiredActions'], str) else [],
                    file_name=row['FileName']
                )
                self.task_details[task_id] = task_detail

        except Exception as e:
            QMessageBox.warning(
                self, "Error", f"Error loading task details: {str(e)}")

    def get_screen_size(self):
        """Get the screen size of the primary display"""
        desktop = QDesktopWidget()
        screen_rect = desktop.availableGeometry(desktop.primaryScreen())
        return screen_rect

    def center_window(self):
        """Center the window on the screen"""
        frame_geometry = self.frameGeometry()
        screen_center = QDesktopWidget().availableGeometry().center()
        frame_geometry.moveCenter(screen_center)
        self.move(frame_geometry.topLeft())

    def create_save_folder(self):
        """Create folder structure for saving files"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            save_folder = os.path.join(current_dir, "Completed_Tasks")
            
            # Create main save folder if it doesn't exist
            if not os.path.exists(save_folder):
                os.makedirs(save_folder)
                
            # Create task-specific folders
            for i in range(1, self.total_tasks + 1):
                task_folder = os.path.join(save_folder, f"Task_{i}")
                if not os.path.exists(task_folder):
                    os.makedirs(task_folder)
                    
            return save_folder
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error creating save folders: {str(e)}")
            return None

    def save_current_document(self):
        """Save the current document"""
        try:
            if not self.word_app or not self.source_doc:
                return False
                
            # Create save folders if they don't exist
            if not hasattr(self, 'save_folder'):
                self.save_folder = self.create_save_folder()
                
            if not self.save_folder:
                return False
                
            # Create save path
            task_folder = os.path.join(self.save_folder, f"Task_{self.current_task}")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Task_{self.current_task}_{timestamp}.docx"
            save_path = os.path.join(task_folder, filename)
            
            # Save document
            self.source_doc.SaveAs(str(save_path))
            return True
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error saving document: {str(e)}")
            return False

    def closeEvent(self, event):
        """Handle application close event"""
        try:
            if self.word_app:
                # Check if there are unsaved changes
                if self.source_doc and not self.source_doc.Saved:
                    reply = QMessageBox.question(
                        self, 
                        'Save Changes?',
                        'Do you want to save your changes before closing?',
                        QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
                    )
                    
                    if reply == QMessageBox.Cancel:
                        event.ignore()
                        return
                    elif reply == QMessageBox.Yes:
                        if self.save_current_document():
                            self.show_save_summary()
                        
                # Close Word application
                try:
                    if self.source_doc:
                        self.source_doc.Close(SaveChanges=False)
                    self.word_app.Quit()
                except:
                    pass
        except:
            pass
            
        event.accept()

    def show_save_summary(self):
        """Show summary of saved files"""
        try:
            if not hasattr(self, 'save_folder'):
                return
                
            # Get list of saved files
            all_files = []
            for i in range(1, self.total_tasks + 1):
                task_folder = os.path.join(self.save_folder, f"Task_{i}")
                if os.path.exists(task_folder):
                    files = os.listdir(task_folder)
                    if files:
                        all_files.extend([os.path.join(f"Task_{i}", f) for f in files])
            
            if not all_files:
                return
                
            # Create summary message
            summary = (
                f"Files have been saved to:\n"
                f"{self.save_folder}\n\n"
                f"Total files saved: {len(all_files)}\n\n"
                "Would you like to open the save folder?"
            )
            
            reply = QMessageBox.question(
                self, 
                'Save Summary', 
                summary,
                QMessageBox.Yes | QMessageBox.No
            )
                
            if reply == QMessageBox.Yes:
                os.startfile(self.save_folder)
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error showing summary: {str(e)}")

    def moveEvent(self, event):
        """Handle window move events to maintain bottom position"""
        super().moveEvent(event)
        new_pos = event.pos()
        
        # If window is moved vertically, reset to bottom
        screen_bottom = self.screen_size.height()
        if new_pos.y() + self.height() != screen_bottom:
            self.move(new_pos.x(), screen_bottom - self.height())
        
        # Keep window within screen bounds horizontally
        if new_pos.x() < 0:
            self.move(0, new_pos.y())
        elif new_pos.x() + self.width() > self.screen_size.width():
            self.move(self.screen_size.width() - self.width(), new_pos.y())

    def position_window_bottom(self):
        """Position window at the bottom of the screen"""
        screen_bottom = self.screen_size.height()
        self.move(self.window_x, screen_bottom - self.height())

    
    def update_layout_for_resize(self):
        """Update layout elements based on window size"""
        current_width = self.width()
        
        # Update button widths
        nav_button_width = int(current_width * 0.1)  # 10% of window width
        for btn in [self.prev_btn, self.mark_complete_btn, 
                   self.mark_review_btn, self.next_btn]:
            if btn:
                btn.setFixedWidth(nav_button_width)
        
        # Update task button widths
        task_button_width = int(current_width * 0.08)  # 8% of window width
        for btn in self.task_buttons:
            if btn:
                btn.setFixedWidth(task_button_width)
        
        # Update progress bar width
        progress_bar = self.findChild(QProgressBar)
        if progress_bar:
            progress_bar.setFixedWidth(int(current_width * 0.15))
        
        # Update description text minimum width
        if hasattr(self, 'description_text'):
            self.description_text.setMinimumWidth(int(current_width * 0.95))

    def resizeEvent(self, event):
        """Handle window resize events"""
        super().resizeEvent(event)
        
        # Don't allow height to be less than minimum
        if self.height() < self.min_height:
            self.resize(self.width(), self.min_height)
        
        # Don't allow width to be less than minimum
        if self.width() < self.min_width:
            self.resize(self.min_width, self.height())
        
        # Update resize grip position
        if hasattr(self, 'resize_grip'):
            self.resize_grip.move(
                self.width() - self.resize_grip.width(),
                self.height() - self.resize_grip.height()
            )
        
        # Update layout elements
        self.update_layout_for_resize()
        
        # Maintain bottom position
        self.position_window_bottom()
        
        # Ensure Word window remains visible
        if self.word_app:
            try:
                self.word_app.Activate()
            except:
                pass

    def moveEvent(self, event):
        """Handle window move events to maintain bottom position"""
        super().moveEvent(event)
        new_pos = event.pos()
        
        # Always keep window at screen bottom while allowing horizontal movement
        screen_bottom = self.screen_size.height()
        self.move(new_pos.x(), screen_bottom - self.height())
        
        # Keep window within screen bounds horizontally
        if new_pos.x() < 0:
            self.move(0, screen_bottom - self.height())
        elif new_pos.x() + self.width() > self.screen_size.width():
            self.move(self.screen_size.width() - self.width(), 
                     screen_bottom - self.height())
            
    def init_resize_grip(self):
        """Initialize the resize grip"""
        self.resize_grip = QSizeGrip(self)
        self.resize_grip.setStyleSheet('''
            QSizeGrip {
                background-color: rgba(255, 255, 255, 0.5);
                width: 16px;
                height: 16px;
                margin: 2px;
                border-radius: 8px;
            }
            QSizeGrip:hover {
                background-color: rgba(255, 255, 255, 0.7);
            }
        ''')

    def initUI(self):
        # Set window properties with resizing allowed
        self.setWindowTitle('Word Associate 2019/365 Skill Review 1 (Q.57)')
        self.setMinimumSize(self.min_width, self.min_height)  # Set minimum size
        self.resize(self.window_width, self.window_height)    # Set initial size
        
        # Set window flags to allow resizing
        self.setWindowFlags(
            Qt.Window |
            Qt.WindowTitleHint |
            Qt.CustomizeWindowHint |
            Qt.WindowCloseButtonHint |
            Qt.WindowMinimizeButtonHint
        )
        
        # Enable resizing
        self.init_resize_grip()
        
        # Set window background style
        self.setStyleSheet('''
            QMainWindow {
                background-color: #2b579a;
                color: white;
                border: 1px solid #1e3f7a;
                border-top: 2px solid #3a6ab0;
            }
        ''')

        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(5)  # Reduced spacing
        main_layout.setContentsMargins(10, 5, 10, 5)  # Reduced margins

        # Create top bar with scaled dimensions
        top_bar = QHBoxLayout()
        top_bar.setContentsMargins(5, 0, 5, 0)

        # # Create pin button for always-on-top toggle
        # self.pin_button = QPushButton()
        # self.pin_button.setFixedSize(24, 24)  # Smaller pin button
        # self.pin_button.clicked.connect(self.toggle_always_on_top)
        # self.update_pin_button_style()
        
        # progress_bar = QProgressBar()
        # progress_bar.setValue(0)
        # progress_bar.setFixedWidth(int(self.window_width * 0.15))
        # progress_bar.setFixedHeight(24)  # Match button height
        # progress_bar.setStyleSheet('''
        #     QProgressBar {
        #         border: 1px solid white;
        #         border-radius: 2px;
        #         text-align: center;
        #         background-color: transparent;
        #     }
        #     QProgressBar::chunk {
        #         background-color: white;
        #     }
        # ''')

        # percentage_label = QLabel('100%')
        # self.timer_label = QLabel('00:00:00')
        # self.timer_label.setStyleSheet('font-family: monospace; font-size: 12px;')
        # # Create pin button for always-on-top toggle
        # self.pin_button = QPushButton()
        # self.pin_button.setFixedSize(24, 24)
        # self.pin_button.clicked.connect(self.toggle_always_on_top)
        # self.update_pin_button_style()

        # Create top bar with controls
        top_bar = QHBoxLayout()
        top_bar.setContentsMargins(5, 0, 5, 0)

        # Create project selection combo box
        self.project_combo = QComboBox()
        self.project_combo.addItem('Project 1 - Word Associate 2019/365')
        self.project_combo.addItem('Project 2 - Word Associate 2019/365')
        self.project_combo.addItem('Project 3 - Word Associate 2019/365')
        self.project_combo.setFixedWidth(250)
        self.project_combo.setStyleSheet('''
            QComboBox {
                background-color: white;
                color: black;
                padding: 3px 10px;
                border-radius: 2px;
                border: none;
            }
            QComboBox:hover {
                background-color: #e0e0e0;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
        ''')
        self.project_combo.currentIndexChanged.connect(self.on_project_changed)

        # Create button container
        button_container = QHBoxLayout()
        button_container.setSpacing(5)

        # Create Pin, Submit, Restart, and End buttons
        self.pin_button = QPushButton('📌')
        self.submit_btn = QPushButton('Submit')
        self.restart_btn = QPushButton('Restart')
        self.end_btn = QPushButton('End')
        
        for btn in [self.pin_button, self.submit_btn, self.restart_btn, self.end_btn]:
            btn.setStyleSheet('''
                QPushButton {
                    background-color: white;
                    color: black;
                    border: none;
                    padding: 3px 10px;
                    border-radius: 2px;
                    font-size: 12px;
                    min-width: 60px;
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                }
            ''')
            btn.setFixedHeight(24)

        self.pin_button.setFixedWidth(30)
        self.pin_button.clicked.connect(self.toggle_always_on_top)
        self.submit_btn.clicked.connect(self.submit_project)
        self.restart_btn.clicked.connect(self.restart_project)
        self.end_btn.clicked.connect(self.end_project)

        # Progress section
        progress_bar = QProgressBar()
        progress_bar.setValue(0)
        progress_bar.setFixedWidth(int(self.window_width * 0.15))
        progress_bar.setFixedHeight(24)
        progress_bar.setStyleSheet('''
            QProgressBar {
                border: 1px solid white;
                border-radius: 2px;
                text-align: center;
                background-color: transparent;
            }
            QProgressBar::chunk {
                background-color: white;
            }
        ''')

        # Create zoom control
        zoom_control = self.create_zoom_control()

        # Timer label
        self.timer_label = QLabel('00:00:00')
        self.timer_label.setStyleSheet('font-family: monospace; font-size: 12px; color: white;')

        # Add all elements to top bar
        top_bar.addWidget(self.pin_button)
        top_bar.addWidget(self.project_combo)
        top_bar.addWidget(self.submit_btn)
        top_bar.addWidget(self.restart_btn)
        top_bar.addWidget(self.end_btn)
        top_bar.addWidget(progress_bar)
        top_bar.addLayout(zoom_control)
        top_bar.addStretch()
        top_bar.addWidget(self.timer_label)

        # Create task bar
        task_bar = QHBoxLayout()
        task_bar.setSpacing(2)  # Reduced spacing between buttons

        # Add task buttons with reduced size
        self.task_buttons = []
        button_width = int(self.window_width * 0.08)  # Slightly wider buttons
        for i in range(1, self.total_tasks + 1):
            task_btn = QPushButton(f'Task {i}')
            task_btn.setFixedHeight(24)  # Match other button heights
            task_btn.setFixedWidth(button_width)
            task_btn.setStyleSheet(self.get_task_button_style())
            task_btn.clicked.connect(lambda checked, num=i: self.go_to_task(num))
            self.task_buttons.append(task_btn)
            task_bar.addWidget(task_btn)

        # Create content area
        content_layout = QHBoxLayout()
        content_layout.setSpacing(5)

        # Create description panel
        description_container = QFrame()
        description_container.setStyleSheet('''
            QFrame {
                background-color: white;
                border-radius: 3px;
                margin: 0;
            }
        ''')
        description_layout = QVBoxLayout(description_container)
        description_layout.setContentsMargins(10, 5, 10, 5)  # Reduced margins

        # Task description text edit
        self.description_text = QTextEdit()
        self.description_text.setStyleSheet('''
            QTextEdit {
                background-color: transparent;
                color: black;
                border: none;
                font-size: 12px;
                line-height: 1.4;
            }
            QTextEdit:focus {
                border: none;
                outline: none;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 8px;
                margin: 0;
            }
            QScrollBar::handle:vertical {
                background: #2b579a;
                min-height: 20px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #1e3f7a;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
                background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
        ''')
        self.description_text.setReadOnly(True)
        self.description_text.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.description_text.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        description_layout.addWidget(self.description_text)

        # Create bottom bar with navigation buttons
        bottom_bar = QHBoxLayout()
        bottom_bar.setContentsMargins(5, 0, 5, 0)

        nav_button_width = int(self.window_width * 0.1)
        self.prev_btn = QPushButton('Previous')
        self.mark_complete_btn = QPushButton('Mark Complete')
        self.mark_review_btn = QPushButton('Mark Review')
        self.next_btn = QPushButton('Next')
        help_btn = QPushButton('Help')

        # Style and connect navigation buttons
        nav_button_style = '''
            QPushButton {
                background-color: white;
                color: black;
                border: none;
                padding: 3px 10px;
                border-radius: 2px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        '''

        for btn in [self.prev_btn, self.mark_complete_btn, self.mark_review_btn, 
                self.next_btn, help_btn]:
            btn.setFixedHeight(24)
            btn.setFixedWidth(nav_button_width)
            btn.setStyleSheet(nav_button_style)
        self.prev_btn.clicked.connect(self.go_to_previous)
        self.next_btn.clicked.connect(self.go_to_next)
        self.mark_complete_btn.clicked.connect(self.mark_task_complete)
        self.mark_review_btn.clicked.connect(self.mark_task_for_review)

        bottom_bar.addWidget(self.prev_btn)
        bottom_bar.addWidget(self.mark_complete_btn)
        bottom_bar.addWidget(self.mark_review_btn)
        bottom_bar.addWidget(self.next_btn)
        bottom_bar.addStretch()
        bottom_bar.addWidget(help_btn)

        # Add all layouts to main layout
        main_layout.addLayout(top_bar)
        main_layout.addLayout(task_bar)
        main_layout.addWidget(description_container, 1)
        main_layout.addLayout(bottom_bar)

        # Initialize UI state
        self.update_navigation_buttons()
        self.update_task_buttons()
        self.update_task_description(self.current_task)

        # Start timer and launch Word
        self.seconds = 0
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_timer)
        self.timer.start(1000)
        self.launch_word()
        # Update resize grip position at the end
        self.resize_grip.move(
            self.width() - self.resize_grip.width(),
            self.height() - self.resize_grip.height()
        )
    def end_project(self):
        """Handle ending the current project"""
        reply = QMessageBox.question(
            self,
            'End Project',
            'Are you sure you want to end this project?\nUnsaved changes will be lost.',
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Save final state if needed
            if self.source_doc and not self.source_doc.Saved:
                save_reply = QMessageBox.question(
                    self,
                    'Save Changes',
                    'Do you want to save your changes before ending?',
                    QMessageBox.Yes | QMessageBox.No
                )
                if save_reply == QMessageBox.Yes:
                    self.save_current_document()
            
            # Close Word and return to project selection
            if self.word_app:
                try:
                    if self.source_doc:
                        self.source_doc.Close(SaveChanges=False)
                    self.word_app.Quit()
                except:
                    pass
            
            # Create and show new skill review window
            self.skill_window = SkillReviewWindow()
            self.skill_window.show()
            self.close()
    def on_project_changed(self, index):
        """Handle project selection change"""
        reply = QMessageBox.question(
            self,
            'Change Project',
            'Are you sure you want to change to a different project? Current progress will be reset.',
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Save current project state if needed
            self.save_current_document()
            
            # Reset states
            self.current_task = 1
            self.task_states = {i: 'incomplete' for i in range(1, 11)}
            
            # Load new project
            self.load_project(index)
        else:
            # Revert combo box selection
            self.project_combo.setCurrentIndex(self.current_project_index)

    def load_project(self, project_index):
        """Load a new project"""
        self.current_project_index = project_index
        
        try:
            # Reset UI
            self.update_task_buttons()
            self.update_navigation_buttons()
            
            # Reset progress
            progress_bar = self.findChild(QProgressBar)
            if progress_bar:
                progress_bar.setValue(0)
            
            # Reset timer
            self.seconds = 0
            self.timer_label.setText('00:00:00')
            
            # Load first task of new project
            self.current_task = 1
            self.update_task_description(self.current_task)
            self.open_source_document(self.current_task)
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error loading project: {str(e)}")

    def submit_project(self):
        """Handle project submission"""
        # Check if all tasks are complete
        all_complete = all(state == 'complete' for state in self.task_states.values())
        
        if not all_complete:
            reply = QMessageBox.question(
                self,
                'Incomplete Tasks',
                'Not all tasks are marked as complete. Do you still want to submit?',
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return

        # Save final state
        if self.save_current_document():
            # Show completion message
            QMessageBox.information(
                self,
                'Project Submitted',
                'Project has been submitted successfully!\nMoving to next project...'
            )
            
            # Move to next project
            next_index = (self.project_combo.currentIndex() + 1) % self.project_combo.count()
            self.project_combo.setCurrentIndex(next_index)
        else:
            QMessageBox.warning(self, "Error", "Failed to save project state")

    def restart_project(self):
        """Restart current project"""
        reply = QMessageBox.question(
            self,
            'Restart Project',
            'Are you sure you want to restart this project? All progress will be lost.',
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Reset states
            self.current_task = 1
            self.task_states = {i: 'incomplete' for i in range(1, 11)}
            
            # Reset UI
            self.update_task_buttons()
            self.update_navigation_buttons()
            
            # Reset progress
            progress_bar = self.findChild(QProgressBar)
            if progress_bar:
                progress_bar.setValue(0)
            
            # Reset timer
            self.seconds = 0
            self.timer_label.setText('00:00:00')
            
            # Load first task
            self.update_task_description(self.current_task)
            self.open_source_document(self.current_task)
    def resizeEvent(self, event):
        """Handle window resize events"""
        super().resizeEvent(event)
        
        # Update resize grip position
        if hasattr(self, 'resize_grip'):
            self.resize_grip.move(
                self.width() - self.resize_grip.width(),
                self.height() - self.resize_grip.height()
            )
        
        # Don't allow height to be less than minimum
        if self.height() < self.min_height:
            self.resize(self.width(), self.min_height)
        
        # Don't allow width to be less than minimum
        if self.width() < self.min_width:
            self.resize(self.min_width, self.height())
        
        # Update layout elements
        self.update_layout_for_resize()
        
        # Ensure Word window remains visible
        if self.word_app:
            try:
                self.word_app.Activate()
            except:
                pass

    def moveEvent(self, event):
        """Handle window move events"""
        super().moveEvent(event)
        new_pos = event.pos()
        new_height = self.height()  # Get current height
        
        # Calculate maximum allowed height
        max_height = self.screen_size.height() - new_pos.y()
        
        # Keep window within vertical bounds
        if new_height > max_height:
            self.resize(self.width(), max_height)
        
        # Keep window within horizontal bounds
        if new_pos.x() < 0:
            self.move(0, new_pos.y())
        elif new_pos.x() + self.width() > self.screen_size.width():
            self.move(self.screen_size.width() - self.width(), new_pos.y())

    def update_layout_for_resize(self):
        """Update layout elements based on window size"""
        current_width = self.width()
        current_height = self.height()
        
        # Update button heights
        button_height = max(24, int(current_height * 0.06))  # Minimum 24px, or 6% of height
        
        # Update navigation button widths and heights
        nav_button_width = int(current_width * 0.1)  # 10% of window width
        for btn in [self.prev_btn, self.mark_complete_btn, 
                self.mark_review_btn, self.next_btn]:
            if btn:
                btn.setFixedWidth(nav_button_width)
                btn.setFixedHeight(button_height)
        
        # Update task button dimensions
        task_button_width = int(current_width * 0.08)  # 8% of window width
        for btn in self.task_buttons:
            if btn:
                btn.setFixedWidth(task_button_width)
                btn.setFixedHeight(button_height)
        
        # Update progress bar dimensions
        progress_bar = self.findChild(QProgressBar)
        if progress_bar:
            progress_bar.setFixedWidth(int(current_width * 0.15))
            progress_bar.setFixedHeight(button_height)
        
        # Update description text size
        if hasattr(self, 'description_text'):
            self.description_text.setMinimumWidth(int(current_width * 0.95))
            # Adjust font size based on window height
            current_font_size = self.zoom_slider.value()
            self.update_font_size()  # This will trigger font size update

    def position_window_bottom(self):
        """Position window at the bottom of the screen maintaining current height"""
        current_height = self.height()
        screen_bottom = self.screen_size.height()
        self.move(self.window_x, screen_bottom - current_height)
    def update_task_description(self, task_number):
        """Update the task description panel with enhanced formatting and black text"""
        if task_number in self.task_details:
            task = self.task_details[task_number]
            font_size = self.zoom_slider.value()
            description_text = f"""
            <div style='margin: 5px 0;'>
                <div style='margin: 0 0 8px 0; font-size: {font_size}px; color: black;'>{task.description}</div>
                <p style='color: #2b579a; font-weight: bold; margin: 5px 0; font-size: {font_size}px;'>Required Actions:</p>
                <ul style='margin: 0; padding-left: 20px; font-size: {font_size}px; color: black;'>
            """
            for action in task.required_actions:
                description_text += f"<li style='margin: 2px 0;'>{action.strip()}</li>"
            description_text += "</ul></div>"
            
            self.description_text.setHtml(description_text)
        else:
            default_text = f"""
            <div style='margin: 5px 0; color: black;'>
                <div style='margin: 0 0 8px 0;'>Task details not available.</div>
            </div>
            """
            self.description_text.setHtml(default_text)

    def update_font_size(self):
        """Update the font size of the description text"""
        font_size = self.zoom_slider.value()
        self.description_text.setStyleSheet(f'''
            QTextEdit {{
                background-color: white;
                color: black;  /* Changed to black text */
                border: none;
                font-size: {font_size}px;
                line-height: 1.6;
                selection-background-color: #2b579a;
                selection-color: white;
            }}
            QTextEdit:focus {{
                border: none;
                outline: none;
            }}
            QScrollBar:vertical {{
                border: none;
                background: #f0f0f0;
                width: 10px;
                margin: 0;
            }}
            QScrollBar::handle:vertical {{
                background: #2b579a;
                min-height: 20px;
                border-radius: 5px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: #1e3f7a;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0;
                background: none;
            }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
                background: none;
            }}
        ''')

    def update_task_description(self, task_number):
        """Update the task description panel with enhanced formatting and black text"""
        if task_number in self.task_details:
            task = self.task_details[task_number]
            font_size = self.zoom_slider.value()
            description_text = f"""
            <div style='margin-bottom: 10px;'>
                <h3 style='color: #2b579a; margin: 0 0 10px 0; font-size: {font_size + 2}px;'>Task {task_number}</h3>
                <div style='margin: 0 0 10px 0; font-size: {font_size}px; color: black;'>{task.description}</div>
                <p style='color: #2b579a; font-weight: bold; margin: 5px 0; font-size: {font_size}px;'>Required Actions:</p>
                <ul style='margin: 0; padding-left: 20px; font-size: {font_size}px; color: black;'>
            """
            for action in task.required_actions:
                description_text += f"<li style='margin: 3px 0;'>{action.strip()}</li>"
            description_text += "</ul></div>"
            
            self.description_text.setHtml(description_text)
        else:
            default_text = f"""
            <div style='margin-bottom: 10px; color: black;'>
                <h3 style='color: #2b579a; margin: 0 0 10px 0;'>Task {task_number}</h3>
                <div style='margin: 0 0 10px 0;'>Task details not available.</div>
            </div>
            """
            self.description_text.setHtml(default_text)


    def go_to_task(self, task_number):
        """Go to a specific task number"""
        if 1 <= task_number <= self.total_tasks:
            self.current_task = task_number
            self.update_task_ui()
            self.update_task_description(task_number)
            
            # Only open document if it's a different file
            if task_number in self.task_details:
                task = self.task_details[task_number]
                file_name = task.file_name
                current_dir = os.path.dirname(os.path.abspath(__file__))
                new_path = os.path.join(current_dir, str(file_name))
                
                if not hasattr(self, 'current_doc_path') or self.current_doc_path != new_path:
                    self.open_source_document(task.file_name)


    def update_task_ui(self):
        """Update the UI elements for the current task"""
        # Update navigation buttons
        self.update_navigation_buttons()

        # Update task buttons
        self.update_task_buttons()

        # Update task description
        self.update_task_description(self.current_task)

        # Update window title
        self.setWindowTitle(f'Task {self.current_task} of {self.total_tasks}')

    def load_task_details(self):
        """Load task details from Excel file"""
        try:
            # Get the directory of the script
            current_dir = os.path.dirname(os.path.abspath(__file__))
            excel_path = os.path.join(
                current_dir, "Project1_Requirements.xlsx")

            if not os.path.exists(excel_path):
                QMessageBox.warning(
                    self, "Error", f"Task file not found: {excel_path}")
                # Initialize with default task if file not found
                self.init_default_tasks()
                return

            # Read Excel file
            df = pd.read_excel(excel_path)

            # Process each row into TaskDetail objects
            for _, row in df.iterrows():
                task_id = row['TaskID']
                task_detail = TaskDetail(
                    task_id=task_id,
                    description=row['Description'],
                    required_actions=row['RequiredActions'].split(
                        ';') if isinstance(row['RequiredActions'], str) else [],
                    file_name=row['FileName']
                )
                self.task_details[task_id] = task_detail

        except Exception as e:
            QMessageBox.warning(
                self, "Error", f"Error loading task details: {str(e)}")
            # Initialize with default task if loading fails
            self.init_default_tasks()

    def init_default_tasks(self):
        """Initialize default tasks if Excel file cannot be loaded"""
        for i in range(1, self.total_tasks + 1):
            self.task_details[i] = TaskDetail(
                task_id=i,
                description=f"Default task {i} description",
                required_actions=["Complete the task requirements"],
                file_name="2019_WE_101_Houseboating.docx"
            )

    def verify_task_completion(self, task_number):
        """Verify if the task requirements are met"""
        if task_number not in self.task_details:
            return False

        task = self.task_details[task_number]
        try:
            # This is where you would implement specific checks for each task
            # For example, checking formatting, content, etc.
            # For now, we'll just return True
            return True
        except Exception as e:
            QMessageBox.warning(
                self, "Error", f"Error verifying task: {str(e)}")
            return False

    def mark_task_complete(self):
        """Mark the current task as complete and save work"""
        if self.verify_task_completion(self.current_task):
            # Save current work
            # save_success = self.save_current_document()
            # save_msg = "\nWork has been saved." if save_success else "\nCould not save work."
                
            self.task_states[self.current_task] = 'complete'
            self.update_task_buttons()
            
            # Calculate progress
            completed_tasks = sum(1 for state in self.task_states.values() if state == 'complete')
            progress = (completed_tasks / self.total_tasks) * 100
            
            # Update progress bar
            progress_bar = self.findChild(QProgressBar)
            if progress_bar:
                progress_bar.setValue(int(progress))
            
            # # Show completion message
            # QMessageBox.information(
            #     self, 
            #     "Success", 
            #     f"Task {self.current_task} completed successfully!{save_msg}"
            # )
            
            # Move to next task if not on last task
            if self.current_task < self.total_tasks:
                self.go_to_next()
        # else:
        #     QMessageBox.warning(
        #         self, 
        #         "Incomplete", 
        #         "Please complete all required actions before marking as complete."
        #     )
    def get_transparent_button_style(self):
        return '''
            QPushButton {
                background-color: transparent;
                color: white;
                border: none;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #1e3f7a;
            }
        '''

    def get_white_button_style(self):
        return '''
            QPushButton {
                background-color: white;
                color: black;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        '''

    def get_task_button_style(self, is_current=False, state='incomplete'):
        base_style = '''
            QPushButton {
                background-color: %s;
                color: white;
                border: none;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #1e3f7a;
            }
        '''

        if is_current:
            return base_style % '#1e3f7a'

        colors = {
            'incomplete': 'transparent',
            'complete': '#4CAF50',
            'review': '#FFA500'
        }
        return base_style % colors[state]

    def update_task_description(self, task_number):
        """Update the task description panel with enhanced formatting"""
        if task_number in self.task_details:
            task = self.task_details[task_number]
            description_text = f"""
            <div style='margin-bottom: 10px;'>
                <h3 style='color: #FFD700; margin: 0 0 10px 0;'>Task {task_number}</h3>
                <div style='margin: 0 0 10px 0;'>{task.description}</div>
                <p style='color: #FFD700; font-weight: bold; margin: 5px 0;'>Required Actions:</p>
                <ul style='margin: 0; padding-left: 20px;'>
            """
            for action in task.required_actions:
                description_text += f"<li style='margin: 3px 0;'>{action.strip()}</li>"
            description_text += "</ul></div>"

            # Set the HTML content
            self.description_text.setHtml(description_text)

    def create_context_menu(self, position):
        """Create context menu for the task description"""
        menu = self.description_text.createStandardContextMenu()
        menu.exec_(self.description_text.mapToGlobal(position))

    def setup_shortcuts(self):
        """Setup keyboard shortcuts for copy operation"""
        copy_shortcut = QShortcut(QKeySequence.Copy, self.description_text)
        copy_shortcut.activated.connect(self.copy_selected_text)

    def copy_selected_text(self):
        """Copy selected text to clipboard"""
        cursor = self.description_text.textCursor()
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            QApplication.clipboard().setText(selected_text)

    def update_timer(self):
        self.seconds += 1
        hours = self.seconds // 3600
        minutes = (self.seconds % 3600) // 60
        seconds = self.seconds % 60
        self.timer_label.setText(f'{hours:02d}:{minutes:02d}:{seconds:02d}')

    def go_to_previous(self):
        """Navigate to the previous task"""
        if self.current_task > 1:
            self.current_task -= 1
            self.update_task_ui()
            self.open_source_document(self.current_task)

    def go_to_next(self):
        """Navigate to the next task"""
        if self.current_task < self.total_tasks:
            self.current_task += 1
            self.update_task_ui()
            self.open_source_document(self.current_task)


    def go_to_task(self, task_number):
        """Go to a specific task number"""
        if 1 <= task_number <= self.total_tasks:
            self.current_task = task_number
            self.update_task_ui()
            self.update_task_description(task_number)
            self.open_source_document(task_number)

    def go_to_previous(self):
        """Navigate to the previous task"""
        if self.current_task > 1:
            self.current_task -= 1
            self.update_task_ui()
            self.update_task_description(self.current_task)
            self.open_source_document(self.current_task)

    def go_to_next(self):
        """Navigate to the next task"""
        if self.current_task < self.total_tasks:
            self.current_task += 1
            self.update_task_ui()
            self.update_task_description(self.current_task)
            self.open_source_document(self.current_task)
            # QMessageBox.warning(
            #     self, "Incomplete", "Please complete all required actions before marking as complete.")

    def launch_word(self):
        """Launch Microsoft Word application"""
        try:
            # Create Word application instance
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = True

            # Set initial file
            self.open_source_document(1)  # Open first task's document

        except Exception as e:
            QMessageBox.warning(
                self, "Error", f"Error launching Word: {str(e)}")

    def closeEvent(self, event):
        """Handle application close event"""
        try:
            # Close Word when application closes
            if self.word_app:
                if self.source_doc:
                    try:
                        self.source_doc.Close(SaveChanges=False)
                    except:
                        pass
                self.word_app.Quit()
        except:
            pass
        event.accept()



        # Calculate progress
        completed_tasks = sum(
            1 for state in self.task_states.values() if state == 'complete')
        progress = (completed_tasks / self.total_tasks) * 100

        # Update progress bar if it exists
        progress_bar = self.findChild(QProgressBar)
        if progress_bar:
            progress_bar.setValue(int(progress))

        # Automatically move to next task if not on last task
        if self.current_task < self.total_tasks:
            self.go_to_next()

    def mark_task_for_review(self):
        """Mark the current task for review"""
        self.task_states[self.current_task] = 'review'
        self.update_task_buttons()

    def update_task_ui(self):
        """Update the UI elements for the current task"""
        # Update navigation buttons
        self.update_navigation_buttons()

        # Update task buttons
        self.update_task_buttons()

        # Update window title
        self.setWindowTitle(f'Task {self.current_task} of {self.total_tasks}')

    def update_navigation_buttons(self):
        """Update the state and text of navigation buttons"""
        # Enable/disable Previous button
        self.prev_btn.setEnabled(self.current_task > 1)

        # Enable/disable Next button
        self.next_btn.setEnabled(self.current_task < self.total_tasks)

        # Update button texts
        if self.current_task == 1:
            self.prev_btn.setText("Previous")
        else:
            self.prev_btn.setText(f"Previous (Task {self.current_task - 1})")

        if self.current_task == self.total_tasks:
            self.next_btn.setText("Next")
        else:
            self.next_btn.setText(f"Next (Task {self.current_task + 1})")

    def update_task_buttons(self):
        """Update the appearance of task buttons based on their states"""
        for i, btn in enumerate(self.task_buttons):
            task_num = i + 1
            is_current = task_num == self.current_task
            state = self.task_states[task_num]
            btn.setStyleSheet(self.get_task_button_style(is_current, state))

    def open_source_document(self, task_number):
        """Open the source document for the given task number"""
        try:
            # Get the directory of the script
            current_dir = os.path.dirname(os.path.abspath(__file__))

            # Construct the source document path
            source_path = os.path.join(
                current_dir, "2019_WE_101_Houseboating.docx")

            if not os.path.exists(source_path):
                QMessageBox.warning(
                    self, "Error", f"Source file not found: {source_path}")
                return

            # Close any existing documents
            if self.source_doc:
                try:
                    self.source_doc.Close()
                except:
                    pass

            # Open the source document
            self.source_doc = self.word_app.Documents.Open(source_path)
            self.word_app.Visible = True

            # Activate Word window
            self.word_app.Activate()

        except Exception as e:
            QMessageBox.warning(
                self, "Error", f"Error opening document: {str(e)}")

    def open_source_document(self, task_number_or_filename):
        """Open the source document for the given task or filename"""
        try:
            # Get the directory of the script
            current_dir = os.path.dirname(os.path.abspath(__file__))
            
            # Determine the file name based on input type
            if isinstance(task_number_or_filename, int):
                if task_number_or_filename in self.task_details:
                    file_name = self.task_details[task_number_or_filename].file_name
                else:
                    file_name = "2019_WE_101_Houseboating.docx"  # Default file
            else:
                file_name = task_number_or_filename
            
            # Track current document path
            new_path = os.path.join(current_dir, str(file_name))
            
            # Check if we already have this document open
            if hasattr(self, 'current_doc_path') and self.current_doc_path == new_path:
                # Document already open, just activate it
                if self.source_doc:
                    self.word_app.Activate()
                    return
            
            # If path is different or no document is open, proceed with opening
            if not os.path.exists(new_path):
                QMessageBox.warning(self, "Error", f"Source file not found: {new_path}")
                return
                
            # Close existing document if it's different
            if self.source_doc:
                try:
                    self.source_doc.Close(SaveChanges=False)
                except:
                    pass
                    
            # Open the new document
            self.source_doc = self.word_app.Documents.Open(str(new_path))
            self.current_doc_path = new_path  # Track current document path
            self.word_app.Visible = True
            self.word_app.Activate()
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error opening document: {str(e)}")

    def verify_task_completion(self, task_number):
        """Verify if the task requirements are met"""
        if task_number not in self.task_details:
            return False

        task = self.task_details[task_number]
        try:
            # This is where you would implement specific checks for each task
            # For example, checking formatting, content, etc.
            # For now, we'll just return True
            return True
        except Exception as e:
            QMessageBox.warning(
                self, "Error", f"Error verifying task: {str(e)}")
            return False

    def launch_word(self):
        """Launch Microsoft Word application"""
        try:
            # Create Word application instance
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = True

        except Exception as e:
            QMessageBox.warning(
                self, "Error", f"Error launching Word: {str(e)}")

    def closeEvent(self, event):
        """Handle application close event"""
        # Close Word when application closes
        if self.word_app:
            try:
                # Close any open documents
                for doc in self.word_app.Documents:
                    doc.Close(SaveChanges=False)
                self.word_app.Quit()
            except:
                pass
        event.accept()
    # ... (rest of the TestWindow methods from your code)

# ... (SkillReviewWindow and MOSTestApp classes remain the same)


class SkillReviewWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('MO-100: Microsoft Word (Office 2019/365)')
        self.setFixedSize(400, 200)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Add title
        title_label = QLabel('CHỌN MỘT BÀI ĐỀ ÔN THI')
        title_label.setStyleSheet('font-weight: bold;')
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # Add combobox
        combo = QComboBox()
        combo.addItem('Word Associate 2019/365 Skill Review 1 (Q.57)')
        combo.addItem('Word Associate 2019/365 Skill Review 2 (Q.50)')
        layout.addWidget(combo)

        # Add buttons
        btn_layout = QHBoxLayout()

        self.retry_btn = QPushButton('Quay lại')
        self.confirm_btn = QPushButton('Xác nhận')

        btn_layout.addWidget(self.retry_btn)
        btn_layout.addWidget(self.confirm_btn)

        layout.addLayout(btn_layout)

        # Connect buttons
        self.retry_btn.clicked.connect(self.show_main_window)
        self.confirm_btn.clicked.connect(self.launch_test)

    def show_main_window(self):
        self.main_window = MOSTestApp()
        self.main_window.show()
        self.close()

    def launch_test(self):
        self.test_window = TestWindow()
        self.test_window.show()
        self.close()


class MOSTestApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Set window properties
        self.setWindowTitle('PHẦN MỀM LUYỆN THI MOS')
        self.setFixedSize(400, 500)

        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setAlignment(Qt.AlignCenter)

        # Add title
        title_label = QLabel('PHẦN MỀM LUYỆN THI MOS')
        title_label.setStyleSheet(
            'color: red; font-size: 16px; font-weight: bold;')
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # Add TIN HOC MOS icon/label
        icon_frame = QFrame()
        icon_frame.setStyleSheet('''
            QFrame {
                background-color: #2b579a;
                border-radius: 10px;
                min-height: 100px;
            }
        ''')
        icon_layout = QVBoxLayout(icon_frame)

        icon_label = QLabel('TIN HOC\nMOS')
        icon_label.setStyleSheet(
            'color: white; font-size: 24px; font-weight: bold;')
        icon_label.setAlignment(Qt.AlignCenter)
        icon_layout.addWidget(icon_label)

        layout.addWidget(icon_frame)

        # Add "CHON MON OFFICE 2019/365" label
        office_label = QLabel('CHỌN MÔN OFFICE 2019/365')
        office_label.setStyleSheet(
            'color: red; font-size: 14px; font-weight: bold;')
        office_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(office_label)

        # Add combobox for selecting Microsoft Word
        self.combo = QComboBox()
        self.combo.addItem('MO-100: Microsoft Word (Office 2019/365)')
        self.combo.setStyleSheet('''
            QComboBox {
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 3px;
            }
        ''')
        layout.addWidget(self.combo)

        # Add buttons
        self.login_btn = QPushButton('Đăng nhập')
        self.login_btn.setStyleSheet('''
            QPushButton {
                background-color: #2b579a;
                color: white;
                padding: 8px;
                border-radius: 4px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #1e3f7a;
            }
        ''')
        layout.addWidget(self.login_btn)

        # Connect login button
        self.login_btn.clicked.connect(self.show_skill_review)

    def show_skill_review(self):
        self.skill_window = SkillReviewWindow()
        self.skill_window.show()
        self.close()
# import pandas as pd
# import os

# # Define the tasks data
# tasks_data = {
#     'TaskID': [
#         1,
#         2,
#         3,
#         4,
#         5,
#         6,
#         7,
#         8,
#         9,
#         10
#     ],
#     'Description': [
#         'Format the document title "FUSION TOMO BUSINESS PLAN" with appropriate styling',
#         'Format the "EXECUTIVE SUMMARY" heading with the provided theme color',
#         'Create a table to display seasonal pricing for houseboats',
#         'Apply appropriate formatting to contact information section',
#         'Create and format a bulleted list for Types of Lodging',
#         'Add page headers and footers to the document',
#         'Insert and format the pricing image',
#         'Format the pricing information with proper currency formatting',
#         'Create and format the numbered objectives list',
#         'Apply styles to all document headings'
#     ],
#     'RequiredActions': [
#         'Apply Bold;Center align;Set font size to 16pt;Apply Gold theme color',
#         'Apply theme color;Left align;Apply Heading 1 style',
#         'Insert table;Apply borders;Merge header cells;Center align headers',
#         'Left align;Add spacing after;Apply normal style',
#         'Create bulleted list;Apply consistent spacing;Set proper indentation',
#         'Add page numbers;Add company name;Set different first page',
#         'Insert image;Set width to 3.5 inches;Apply text wrapping',
#         'Apply currency format;Add spacing between sections;Align decimals',
#         'Create numbered list;Apply consistent indentation;Set proper spacing',
#         'Apply Heading styles;Ensure consistent formatting;Update table of contents'
#     ],
#     'FileName': [
#         '2019_WE_101_Houseboating.docx',
#         '2019_WE_101_Houseboating.docx',
#         '2019_WE_101_Houseboating.docx',
#         '2019_WE_101_Houseboating.docx',
#         '2019_WE_101_Houseboating.docx',
#         '2019_WE_101_Houseboating.docx',
#         '2019_WE_101_Houseboating.docx',
#         '2019_WE_101_Houseboating.docx',
#         '2019_WE_101_Houseboating.docx',
#         '2019_WE_101_Houseboating.docx'
#     ]
# }

# # Create DataFrame
# df = pd.DataFrame(tasks_data)

# # Get the current directory
# current_dir = os.path.dirname(os.path.abspath(__file__))

# # Save to Excel file
# excel_path = os.path.join(current_dir, "Project1_Requirements.xlsx")
# df.to_excel(excel_path, index=False, sheet_name='Tasks')


# # Read back and display the first few rows to verify
# print("Excel file created successfully at:", excel_path)
# print("\nPreview of the created file:")
# df_verify = pd.read_excel(excel_path)
# print(df_verify.head())
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MOSTestApp()
    ex.show()
    sys.exit(app.exec_())
