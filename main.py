"""
Aghosh Care System - Login Module
Entry point for the application. Handles user authentication,
registration, and launches the main dashboard.
"""
from Splash import SplashScreen
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import os
import sys

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from datetime import datetime

from db_helper import get_connection, initialize_database, set_window_icon
from styles import (COLORS, FONTS, apply_theme, center_window,
                    setup_modal_window, create_window_header, styled_entry,
                    ORG_SUFFIX)


# ============================================================
# DATABASE HELPERS
# ============================================================

def get_last_logged_in_username():
    """Retrieve the username of the last logged-in user."""
    conn = get_connection()
    if not conn:
        return ""
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT Username
            FROM tblUsers
            WHERE IsLoggedIn = 1
            ORDER BY UserID DESC
            LIMIT 1
        """)
        row = cursor.fetchone()
        conn.close()
        return row.Username if row else ""
    except Exception:
        return ""


# ============================================================
# LOGIN HANDLER
# ============================================================

def on_login():
    """Validate credentials and launch the dashboard on success."""
    username = username_entry.get()
    password = password_entry.get()

    if not username or not password:
        messagebox.showwarning("Input Error", "Please enter both username and password.")
        return

    conn = get_connection()
    if conn:
        try:
            cursor = conn.cursor()

            cursor.execute("SELECT COUNT(*) AS cnt FROM tblUsers")
            user_count = cursor.fetchone()[0]

            if user_count == 0:
                pass
            else:
                cursor.execute(
                    "SELECT * FROM tblUsers WHERE Username = ? AND PasswordHash = ?",
                    (username, password)
                )
                row = cursor.fetchone()

                if row:
                    try:
                        cursor.execute("UPDATE tblUsers SET IsLoggedIn = 0")
                        cursor.execute(
                            "UPDATE tblUsers SET IsLoggedIn = 1 WHERE Username = ?",
                            (username,)
                        )
                        conn.commit()

                        logged_in_role = row.Role
                        login_button.config(state="disabled")

                        root.progress.pack(pady=(10, 5))
                        root.progress_label.pack()

                        start_dashboard_loading(username, row.Role, row.CenterID)

                    except Exception as e:
                        conn.rollback()
                        messagebox.showerror("Database Error", str(e))
                else:
                    messagebox.showwarning("Login Failed", "Invalid username or password.")
        except Exception as e:
            messagebox.showerror("Error", f"Login failed: {e}")
    else:
        messagebox.showerror("Connection Failed", "Unable to connect to the database.")


def start_dashboard_loading(username, role, center_id):
    """Animate the progress bar with real status messages then launch the dashboard."""
    root.progress_var.set(0)

    steps = [
        (10, "Connecting to database..."),
        (25, "Loading user session..."),
        (40, "Preparing modules..."),
        (55, "Loading child records..."),
        (70, "Initializing dashboard..."),
        (85, "Importing dashboard module..."),
        (95, "Almost ready..."),
    ]

    step_index = [0]

    def step():
        if step_index[0] < len(steps):
            val, msg = steps[step_index[0]]
            root.progress_var.set(val)
            root.progress_label.config(text=msg)
            step_index[0] += 1
            root.after(200, step)
        else:
            root.progress_var.set(100)
            root.progress_label.config(text="Launching dashboard...")
            root.after(100, lambda: launch_dashboard(username, role, center_id))

    step()


def launch_dashboard(username, role, center_id):
    """Import and open the main dashboard window."""
    try:
        import dashboard
        root.destroy()
        dashboard.open_dashboard(username, role, center_id=center_id)
    except Exception as e:
        messagebox.showerror("Error", str(e))


# ============================================================
# ADMIN AUTHENTICATION WINDOW
# ============================================================

def open_admin_authentication_window():
    """Open a modal dialog requiring admin credentials before registration."""
    auth_window = tk.Toplevel(root)
    setup_modal_window(auth_window, root, "Admin Authentication", 450, 300)
    auth_window.configure(bg=COLORS["bg"])

    # -- Header --
    create_window_header(auth_window, "Admin Authentication", "Verify your identity")

    # -- Form --
    form_frame = tk.Frame(auth_window, bg=COLORS["bg"])
    form_frame.pack(padx=40, pady=30)

    tk.Label(form_frame, text="Username:", font=FONTS["body"],
             bg=COLORS["bg"]).grid(row=0, column=0, sticky="e", pady=8, padx=(0, 10))
    auth_username = styled_entry(form_frame, width=25)
    auth_username.grid(row=0, column=1, pady=8)
    auth_username.focus_set()

    tk.Label(form_frame, text="Password:", font=FONTS["body"],
             bg=COLORS["bg"]).grid(row=1, column=0, sticky="e", pady=8, padx=(0, 10))
    auth_password = styled_entry(form_frame, width=25)
    auth_password.config(show="*")
    auth_password.grid(row=1, column=1, pady=8)

    def on_authenticate():
        uname = auth_username.get()
        pwd = auth_password.get()

        if not uname or not pwd:
            messagebox.showwarning("Input Error", "Enter both fields.")
            return

        conn = get_connection()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT Role FROM tblUsers WHERE Username = ? AND PasswordHash = ?",
                    (uname, pwd)
                )
                result = cursor.fetchone()
                conn.close()

                if result and result.Role.lower() == "admin":
                    auth_window.destroy()
                    open_register_window()
                else:
                    messagebox.showerror("Access Denied",
                                         "You are not authorized to register users.")
            except Exception as e:
                messagebox.showerror("Error", f"Authentication failed: {e}")
        else:
            messagebox.showerror("Connection Failed",
                                 "Unable to connect to the database.")

    auth_window.bind('<Return>', lambda event: on_authenticate())

    # -- Button --
    btn_frame = tk.Frame(auth_window, bg=COLORS["bg"])
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="Authenticate", command=on_authenticate,
               style="Primary.TButton").pack()


# ============================================================
# REGISTER WINDOW
# ============================================================

def open_register_window():
    """Open a modal dialog for registering a new user."""
    reg_window = tk.Toplevel(root)
    setup_modal_window(reg_window, root, "Register New User", 480, 420)
    reg_window.configure(bg=COLORS["bg"])

    # -- Header --
    create_window_header(reg_window, "Register New User", "Create a new user account")

    # -- Form --
    form_frame = tk.Frame(reg_window, bg=COLORS["bg"])
    form_frame.pack(padx=40, pady=20)

    tk.Label(form_frame, text="Username:", font=FONTS["body"],
             bg=COLORS["bg"]).grid(row=0, column=0, sticky="e", pady=8, padx=(0, 10))
    reg_username = styled_entry(form_frame, width=25)
    reg_username.grid(row=0, column=1, pady=8)
    reg_username.focus_set()

    tk.Label(form_frame, text="Password:", font=FONTS["body"],
             bg=COLORS["bg"]).grid(row=1, column=0, sticky="e", pady=8, padx=(0, 10))
    reg_password = styled_entry(form_frame, width=25)
    reg_password.config(show="*")
    reg_password.grid(row=1, column=1, pady=8)

    tk.Label(form_frame, text="Role:", font=FONTS["body"],
             bg=COLORS["bg"]).grid(row=2, column=0, sticky="e", pady=8, padx=(0, 10))
    role_var = tk.StringVar()
    role_dropdown = ttk.Combobox(form_frame, textvariable=role_var,
                                 values=["Administrator", "Editor", "User"],
                                 state="readonly", font=FONTS["entry"], width=23)
    role_dropdown.grid(row=2, column=1, pady=8)
    role_dropdown.set("User")

    tk.Label(form_frame, text="Center:", font=FONTS["body"],
             bg=COLORS["bg"]).grid(row=3, column=0, sticky="e", pady=8, padx=(0, 10))
    center_var = tk.StringVar()
    center_dropdown = ttk.Combobox(form_frame, textvariable=center_var,
                                    state="readonly", font=FONTS["entry"], width=23)
    center_dropdown.grid(row=3, column=1, pady=8)

    def load_centers():
        conn = get_connection()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT CenterID, CenterName FROM tblCenters")
                rows = cursor.fetchall()
                conn.close()

                center_options = [(row.CenterName, row.CenterID) for row in rows]
                center_dropdown["values"] = [name for name, _id in center_options]

                global center_id_map
                center_id_map = {name: _id for name, _id in center_options}

                if center_options:
                    center_dropdown.set(center_options[0][0])
            except Exception as e:
                messagebox.showerror("Database Error", f"Failed to load centers: {e}")
        else:
            messagebox.showerror("Database Error", "Could not connect to the database.")

    load_centers()

    def on_register():
        username = reg_username.get()
        password = reg_password.get()
        role = role_var.get()
        selected_center_name = center_var.get()

        if not username or not password:
            messagebox.showwarning("Input Error", "Username and Password are required.")
            return

        if not selected_center_name:
            messagebox.showwarning("Input Error", "Please select a center.")
            return

        try:
            center_id = center_id_map[selected_center_name]
        except KeyError:
            messagebox.showerror("Selection Error", "Invalid center selected.")
            return

        conn = get_connection()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO tblUsers
                    (Username, PasswordHash, Role, CenterID)
                    VALUES (?, ?, ?, ?)
                """, (username, password, role, center_id))
                conn.commit()
                conn.close()
                messagebox.showinfo("Success", "User registered successfully!")
                reg_window.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Registration failed: {e}")
        else:
            messagebox.showerror("Connection Failed",
                                 "Unable to connect to the database.")

    reg_window.bind('<Return>', lambda event: on_register())

    # -- Button --
    btn_frame = tk.Frame(reg_window, bg=COLORS["bg"])
    btn_frame.pack(pady=15)
    ttk.Button(btn_frame, text="Register", command=on_register,
               style="Success.TButton").pack()


def on_register_link_click():
    """Handle the 'Register' link click."""
    open_admin_authentication_window()


# ============================================================
# LOGIN UI SETUP
# ============================================================

def setup_login_ui(root_window):
    """Build the professional login interface."""
    global root, username_entry, password_entry, login_button

    root = root_window
    root.title("Aghosh Care System - Login" + ORG_SUFFIX)
    root.configure(bg=COLORS["bg"])
    center_window(root, 480, 520)
    root.resizable(False, False)

    # -- Apply theme --
    apply_theme(root)

    # -- Header banner --
    header = tk.Frame(root, bg=COLORS["bg_header"], height=80)
    header.pack(fill="x")
    header.pack_propagate(False)

    tk.Label(header, text="Aghosh Care System",
             font=FONTS["title"], bg=COLORS["bg_header"],
             fg=COLORS["text_light"]).pack(pady=(15, 2))
    tk.Label(header, text="Orphan Care Management",
             font=FONTS["small"], bg=COLORS["bg_header"],
             fg=COLORS["text_muted"]).pack()

    # -- Login card --
    card_outer = tk.Frame(root, bg=COLORS["bg"])
    card_outer.pack(expand=True, fill="both", padx=40, pady=20)

    card = tk.Frame(card_outer, bg=COLORS["bg_card"], bd=1, relief="solid",
                    highlightbackground=COLORS["border"], highlightthickness=1)
    card.pack(expand=True, fill="both")

    tk.Label(card, text="Sign In", font=FONTS["heading"],
             bg=COLORS["bg_card"], fg=COLORS["primary"]).pack(pady=(25, 15))

    # -- Form fields --
    form_frame = tk.Frame(card, bg=COLORS["bg_card"])
    form_frame.pack(padx=40, pady=5)

    tk.Label(form_frame, text="Username", font=FONTS["small_bold"],
             bg=COLORS["bg_card"], fg=COLORS["text_secondary"],
             anchor="w").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 3))
    username_entry = styled_entry(form_frame, width=30)
    username_entry.grid(row=1, column=0, columnspan=2, pady=(0, 12), ipady=4)

    last_user = get_last_logged_in_username()
    if last_user:
        username_entry.insert(0, last_user)

    tk.Label(form_frame, text="Password", font=FONTS["small_bold"],
             bg=COLORS["bg_card"], fg=COLORS["text_secondary"],
             anchor="w").grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 3))
    password_entry = styled_entry(form_frame, width=30)
    password_entry.config(show="*")
    password_entry.grid(row=3, column=0, columnspan=2, pady=(0, 12), ipady=4)

    # -- Login button --
    login_button = ttk.Button(card, text="Login", command=on_login,
                               style="Primary.TButton", width=20)
    login_button.pack(pady=(10, 5))

    # -- Progress bar (hidden until login) --
    root.progress_var = tk.IntVar(value=0)
    root.progress = ttk.Progressbar(card, variable=root.progress_var,
                                     maximum=100, length=280,
                                     style="Horizontal.TProgressbar")
    root.progress_label = tk.Label(card, text="", font=FONTS["small"],
                                    fg=COLORS["text_secondary"],
                                    bg=COLORS["bg_card"])

    # -- Register link --
    register_link = tk.Label(card, text="Register a New User",
                              fg=COLORS["accent"], cursor="hand2",
                              font=("Segoe UI", 10, "underline"),
                              bg=COLORS["bg_card"])
    register_link.pack(pady=(8, 20))
    register_link.bind("<Button-1>", lambda e: on_register_link_click())

    # -- Keyboard shortcuts --
    root.bind('<Escape>', lambda event: root.destroy())
    root.bind('<Return>', lambda event: on_login())

    username_entry.after(100, lambda: username_entry.select_range(0, tk.END))


# ============================================================
# APPLICATION ENTRY POINTS
# ============================================================

def open_login_window():
    """Open the login window (called after splash or on session expiry)."""
    global root
    setup_login_ui(root)

    root.after(200, lambda: (
        root.deiconify(),
        root.lift()
    ))

    root.mainloop()


def start_login():
    """Entry point alias for login."""
    open_login_window()


# ============================================================
# MAIN ENTRY POINT
# ============================================================

if __name__ == "__main__":
    initialize_database()

    root = tk.Tk()
    root.withdraw()
    set_window_icon(root)

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    PHOTO_DIR = os.path.join(BASE_DIR, "Photo")
    os.makedirs(PHOTO_DIR, exist_ok=True)

    def start_login():
        setup_login_ui(root)
        root.deiconify()
        root.focus_force()
        password_entry.focus_set()

    SplashScreen(root, start_login)
    root.mainloop()
