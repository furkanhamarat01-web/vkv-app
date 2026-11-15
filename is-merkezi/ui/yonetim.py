"""Yönetici ekranları: ekip, talep kodu, kullanıcı, parola isteği ve giriş kaydı."""

import tkinter as tk
from tkinter import messagebox, ttk

import hesap

FONT = ".AppleSystemUIFont"
CARD = "#FFFFFF"
INK = "#14182B"
MUTED = "#6B7287"
LINE = "#E3E6EE"
CANVAS = "#F3F4F7"
ACCENT = "#3250E6"
ACCENT_HOVER = "#2843C4"
ROLES = ("Yönetici", "Sorumlu", "Kullanıcı", "İzleyici")


def render(parent, app, page):
    for child in parent.winfo_children():
        child.destroy()
    parent.pack(fill="both", expand=True)
    try:
        if page == "Ekipler":
            _teams(parent, app)
        elif page == "Talep kodları":
            _codes(parent, app)
        elif page == "Kullanıcılar":
            _users(parent, app)
        else:
            _logs(parent, app)
    except (hesap.ApiError, ValueError) as exc:
        tk.Label(parent, text=str(exc), bg=CANVAS, fg=INK, font=(FONT, 13)).pack(anchor="w", padx=22, pady=20)


def _title(parent, text):
    tk.Label(parent, text=text, bg=CANVAS, fg=INK, font=(FONT, 20, "bold")).pack(anchor="w", padx=22, pady=(8, 8))


def _teams(parent, app):
    _title(parent, "Ekipler")
    name = tk.StringVar()
    topic = tk.StringVar()
    row = tk.Frame(parent, bg=CANVAS)
    row.pack(fill="x", padx=22)
    tk.Entry(row, textvariable=name, width=22, font=(FONT, 13)).pack(side="left", ipady=4)
    tk.Entry(row, textvariable=topic, width=22, font=(FONT, 13)).pack(side="left", padx=8, ipady=4)

    def add():
        try:
            hesap.client().add_team(name.get(), "", topic.get())
        except ValueError as exc:
            messagebox.showwarning(app.title(), str(exc), parent=app)
            return
        render(parent, app, "Ekipler")

    tk.Button(row, text="Ekle", command=add).pack(side="left")
    for team in hesap.client().teams():
        tk.Label(parent, text=f"{team['code']}  ·  {team['name']}  ·  {team.get('topic') or ''}", bg=CANVAS, fg=INK, font=(FONT, 13)).pack(anchor="w", padx=22, pady=2)


def _codes(parent, app):
    _title(parent, "Talep kodları")
    for item in hesap.client().request_codes():
        shared = "  ·  tüm ekipler" if item.get("all_teams") else ""
        tk.Label(
            parent,
            text=f"{item['code']}  ·  {item.get('title') or ''}  ·  {item.get('team_name') or ''}{shared}",
            bg=CANVAS, fg=INK, font=(FONT, 13),
        ).pack(anchor="w", padx=22, pady=2)


def _users(parent, app):
    _title(parent, "Kullanıcılar")
    store = hesap.client()
    resets = store.resets()
    if resets:
        tk.Label(parent, text="Parola istekleri", bg=CANVAS, fg=INK, font=(FONT, 14, "bold")).pack(anchor="w", padx=22, pady=(4, 4))
    for item in resets:
        line = tk.Frame(parent, bg=CANVAS)
        line.pack(fill="x", padx=22, pady=2)
        tk.Label(line, text=f"{item.get('name') or ''}  ·  {item.get('username') or ''}", bg=CANVAS, fg=INK, font=(FONT, 13)).pack(side="left")
        temporary = tk.StringVar()
        tk.Entry(line, textvariable=temporary, show="•", width=16, font=(FONT, 12)).pack(side="left", padx=8)

        def give(reset_id=item["id"], var=temporary):
            if len(var.get()) < 6:
                messagebox.showwarning(app.title(), "Geçici parola en az 6 karakter olmalı.", parent=app)
                return
            try:
                store.grant_reset(reset_id, var.get())
            except ValueError as exc:
                messagebox.showwarning(app.title(), str(exc), parent=app)
                return
            render(parent, app, "Kullanıcılar")

        tk.Button(line, text="Geçici parola ver", command=give).pack(side="left")

    form = tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightthickness=1, padx=12, pady=10)
    form.pack(fill="x", padx=22, pady=12)
    fields = {}
    for label in ("Ad", "Kullanıcı adı", "E-posta", "Parola"):
        tk.Label(form, text=label, bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(anchor="w")
        var = tk.StringVar()
        tk.Entry(form, textvariable=var, show="•" if label == "Parola" else "", font=(FONT, 13)).pack(fill="x", ipady=4, pady=(0, 6))
        fields[label] = var
    role = tk.StringVar(value="Kullanıcı")
    ttk.Combobox(form, textvariable=role, values=list(ROLES), state="readonly").pack(anchor="w", pady=4)

    def add():
        try:
            user_id = store.add_user(
                fields["Ad"].get(), fields["E-posta"].get(), fields["Parola"].get(), "",
                role.get() == "Yönetici", fields["Kullanıcı adı"].get(), role=role.get(),
            )
            teams = store.teams()
            if teams:
                store.add_member(user_id, teams[0]["id"])
        except ValueError as exc:
            messagebox.showwarning(app.title(), str(exc), parent=app)
            return
        render(parent, app, "Kullanıcılar")

    tk.Button(form, text="Kullanıcı ekle", command=add).pack(anchor="e", pady=(8, 0))
    for user in store.users():
        tk.Label(
            parent,
            text=f"{user.get('name') or ''}  ·  {user.get('username') or ''}  ·  {user.get('role') or ''}",
            bg=CANVAS, fg=INK, font=(FONT, 13),
        ).pack(anchor="w", padx=22, pady=2)


def _logs(parent, app):
    _title(parent, "Giriş kayıtları")
    for row in hesap.client().login_logs():
        tk.Label(
            parent,
            text=f"{row.get('created_at') or ''}  ·  {row.get('event') or ''}  ·  {row.get('name') or row.get('email') or ''}",
            bg=CANVAS, fg=INK, font=(FONT, 13),
        ).pack(anchor="w", padx=22, pady=2)
