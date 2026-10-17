# -*- coding: utf-8 -*-
"""
================================================================================
Модул: GitHub Синхронизация и Системна Диагностика
================================================================================
Осигурява двупосочна връзка, управление на версиите и автоматизиран Commit/Push
към профила на потребителя в GitHub през Dulwich (Pure-Python Git) и GitHub API.
"""

import os
import sys
import json
import datetime
import requests
import streamlit as st

parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
root_dir = os.path.dirname(parent_dir)

try:
    import dulwich.porcelain as git
    from dulwich.repo import Repo
    HAS_DULWICH = True
except ImportError:
    HAS_DULWICH = False

CREDENTIALS_FILE = os.path.join(root_dir, ".github_config.json")


def load_github_config():
    """Зарежда запаметената конфигурация за GitHub."""
    if os.path.exists(CREDENTIALS_FILE):
        try:
            with open(CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_github_config(config):
    """Запазва конфигурацията за GitHub локално."""
    try:
        with open(CREDENTIALS_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


def get_git_status():
    """Връща статуса на Git хранилището в работната папка."""
    git_dir = os.path.join(root_dir, ".git")
    if not os.path.exists(git_dir):
        return {"initialized": False}

    try:
        status = git.status(root_dir)
        repo = Repo(root_dir)
        head_commit = repo.head().decode() if repo.head() else "Няма комити"
        
        # Преброяване на промените
        staged_count = len(status.staged.get('add', [])) + len(status.staged.get('delete', [])) + len(status.staged.get('modify', []))
        unstaged_count = len(status.unstaged)
        untracked_count = len(status.untracked)

        return {
            "initialized": True,
            "head": head_commit[:8],
            "staged_count": staged_count,
            "unstaged_count": unstaged_count,
            "untracked_count": untracked_count,
            "untracked": status.untracked[:15],
            "unstaged": status.unstaged[:15]
        }
    except Exception as e:
        return {"initialized": True, "error": str(e)}


def create_github_repo(token, repo_name, private=True):
    """Създава ново хранилище в профила на потребителя в GitHub през REST API."""
    url = "https://api.github.com/user/repos"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }
    payload = {
        "name": repo_name,
        "private": private,
        "description": "Верига магазини „Бакалия“ — Централен софтуерен хъб и Dashboard за търговски и финансов контрол"
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=15)
    return resp.status_code in [200, 201], resp.json()


def sync_to_github(token, username, repo_name, commit_msg):
    """Извършва Add, Commit и Push към GitHub."""
    git_dir = os.path.join(root_dir, ".git")
    if not os.path.exists(git_dir):
        git.init(root_dir)

    # 1. Добавяне на променените файлове
    git.add(root_dir)

    # 2. Commit
    author_str = f"{username} <{username}@users.noreply.github.com>".encode('utf-8')
    msg_bytes = commit_msg.encode('utf-8')
    commit_id = git.commit(root_dir, message=msg_bytes, author=author_str, committer=author_str)

    # 3. Настройване на отдалечено хранилище (remote)
    remote_url = f"https://{username}:{token}@github.com/{username}/{repo_name}.git"
    
    try:
        git.remote_remove(root_dir, "origin")
    except Exception:
        pass
    git.remote_add(root_dir, "origin", remote_url)

    # 4. Push
    git.push(root_dir, remote_url, refspecs=["refs/heads/master:refs/heads/main"])
    return True, commit_id.decode() if isinstance(commit_id, bytes) else str(commit_id)


def render_github_sync():
    """Главен рендер за Таб: GitHub Синхронизация & Диагностика."""
    st.subheader("🐙 Синхронизация с GitHub & Системно състояние")
    st.caption("Автоматизирано архивиране на кода и синхронизация към Вашия акаунт в GitHub")

    if not HAS_DULWICH:
        st.error("Липсва модулът dulwich за Git синхронизация.")
        return

    # Зареждане на съществуващи настройки
    cfg = load_github_config()

    col_cfg1, col_cfg2 = st.columns(2)
    with col_cfg1:
        st.markdown("#### 🔑 Настройки за достъп до GitHub:")
        gh_user = st.text_input("GitHub потребителско име (Username):", value=cfg.get("username", ""), key="gh_user_input")
        gh_repo = st.text_input("Име на хранилището (Repository):", value=cfg.get("repo", "bakaliya-dashboard"), key="gh_repo_input")
        gh_token = st.text_input("GitHub Personal Access Token (PAT):", value=cfg.get("token", ""), type="password", key="gh_token_input", help="Токенът изисква права 'repo'.")

        c_save, c_create = st.columns(2)
        with c_save:
            if st.button("💾 Запази настройките", key="btn_save_gh_cfg"):
                if gh_user and gh_token and gh_repo:
                    save_github_config({"username": gh_user, "repo": gh_repo, "token": gh_token})
                    st.success("✅ Настройките са записани!")
                else:
                    st.warning("Моля попълнете всички полета.")

        with c_create:
            if st.button("🛠️ Създай хранилище в GitHub", key="btn_create_remote_repo"):
                if gh_token and gh_repo:
                    with st.spinner("Създаване на хранилището в GitHub..."):
                        ok, res = create_github_repo(gh_token, gh_repo, private=True)
                        if ok:
                            st.success(f"✅ Хранилището `{gh_repo}` беше успешно създадено в GitHub като Private!")
                        elif res.get("message") == "Repository creation failed." and "name already exists" in str(res):
                            st.info(f"ℹ️ Хранилището `{gh_repo}` вече съществува във Вашия акаунт.")
                        else:
                            st.error(f"Грешка: {res.get('message', res)}")
                else:
                    st.warning("Въведете токен и име на хранилището.")

    with col_cfg2:
        st.markdown("#### 📊 Статус на Git хранилището:")
        git_stat = get_git_status()

        if not git_stat["initialized"]:
            st.warning("Git хранилището все още не е инициализирано в работната папка.")
            if st.button("⚡ Инициализирай Git хранилище сега", key="btn_init_git"):
                git.init(root_dir)
                st.success("✅ Локалното Git хранилище е инициализирано успешно!")
                st.rerun()
        else:
            st.success(f"🟢 Хранилище: Инициализирано | Последен комит: `{git_stat.get('head', '—')}`")
            m1, m2, m3 = st.columns(3)
            m1.metric("Индекс", git_stat.get("staged_count", 0))
            m2.metric("Променени", git_stat.get("unstaged_count", 0))
            m3.metric("Нови файлове", git_stat.get("untracked_count", 0))

            if git_stat.get("untracked"):
                with st.expander("Виж нови неиндексирани файлове"):
                    for u in git_stat["untracked"]:
                        st.caption(f"- {u}")

    st.markdown("---")
    st.markdown("### 🚀 Синхронизиране (Commit & Push):")

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    default_msg = f"Актуализация на Бакалия Dashboard: {now_str}"
    commit_msg = st.text_input("Описание на промените (Commit message):", value=default_msg, key="gh_commit_msg")

    col_sync1, col_sync2 = st.columns([1, 2])
    with col_sync1:
        if st.button("🚀 Синхронизирай с GitHub сега", type="primary", key="btn_push_github"):
            if not gh_token or not gh_user or not gh_repo:
                st.error("❌ Моля, попълнете и запазете Вашите GitHub данни (потребителско име, хранилище и токен).")
            else:
                with st.spinner("Индексиране на файловете, създаване на комит и качване към GitHub..."):
                    try:
                        ok, cid = sync_to_github(gh_token, gh_user, gh_repo, commit_msg)
                        if ok:
                            st.success(f"✅ Успешна синхронизация към GitHub! Комит: `{cid[:8]}`")
                            st.balloons()
                    except Exception as e:
                        st.error(f"❌ Грешка при синхронизация: {e}")

    with col_sync2:
        st.info("🔒 **Сигурност на данните:** Базите данни (`*.FDB`), локалните пароли и логовете са автоматично изключени от качване чрез `.gitignore`.")

    st.markdown("---")
    st.markdown("### 🩺 Диагностика на свързаността:")
    d_c1, d_c2, d_c3 = st.columns(3)

    # 1. CS.FDB
    with d_c1:
        try:
            import firebirdsql
            c_test = firebirdsql.connect(host="192.168.5.254", database=r"E:\DB\CS.FDB", user="SYSDBA", password="masterkey", port=3050)
            c_test.close()
            st.success("🟢 CS.FDB (Централен склад): Свързан")
        except Exception:
            st.error("🔴 CS.FDB: Няма връзка (192.168.5.254)")

    # 2. MONTEVIDEO.FDB
    with d_c2:
        try:
            import firebirdsql
            m_test = firebirdsql.connect(host="192.168.5.254", database=r"E:\db\MONTEVIDEO.FDB", user="SYSDBA", password="masterkey", port=3050)
            m_test.close()
            st.success("🟢 MONTEVIDEO.FDB: Свързан")
        except Exception:
            st.error("🔴 MONTEVIDEO.FDB: Няма връзка")

    # 3. GitHub API
    with d_c3:
        if gh_token:
            try:
                r_test = requests.get("https://api.github.com/user", headers={"Authorization": f"token {gh_token}"}, timeout=5)
                if r_test.status_code == 200:
                    st.success(f"🟢 GitHub API: Оторизиран ({r_test.json().get('login')})")
                else:
                    st.warning(f"🟡 GitHub API: Грешен токен ({r_test.status_code})")
            except Exception:
                st.error("🔴 GitHub API: Недостъпен")
        else:
            st.info("⚪ GitHub API: Не е конфигуриран токен")
