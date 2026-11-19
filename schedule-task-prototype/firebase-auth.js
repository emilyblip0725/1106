import { initializeApp } from "https://www.gstatic.com/firebasejs/10.14.1/firebase-app.js";
import { getAuth, GoogleAuthProvider, signInWithPopup, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.14.1/firebase-auth.js";
import { firebaseConfig } from "./firebase-config.js";

const gate = document.querySelector("#authGate");
const shell = document.querySelector("#appShell");
const error = document.querySelector("#authError");
const login = document.querySelector("#googleSignIn");
const logout = document.querySelector("#signOut");
const hasConfig = firebaseConfig.apiKey !== "YOUR_API_KEY" && firebaseConfig.projectId !== "YOUR_PROJECT_ID" && firebaseConfig.appId !== "YOUR_APP_ID";

function showUser(user) {
  const allowed = Boolean(user);
  document.body.classList.toggle("authenticated", allowed);
  gate.hidden = allowed;
  shell.setAttribute("aria-hidden", String(!allowed));
  if (user) {
    const name = user.displayName || user.email || "已登入";
    document.querySelector("#profileName").textContent = name;
    document.querySelector("#profileAvatar").textContent = (user.displayName || user.email || "U").slice(0, 1).toUpperCase();
    document.querySelector("#topAvatar").textContent = (user.displayName || "U").slice(0, 1).toUpperCase();
  }
}

if (!hasConfig) {
  showUser(null);
  error.textContent = "請先完成 firebase-config.js 設定，再啟用登入。";
  login.disabled = true;
} else {
  const app = initializeApp(firebaseConfig);
  const auth = getAuth(app);
  login.addEventListener("click", async () => {
    error.textContent = "";
    login.disabled = true;
    try { await signInWithPopup(auth, new GoogleAuthProvider()); }
    catch (e) { error.textContent = e.code === "auth/unauthorized-domain" ? "此網域尚未加入 Firebase 授權網域。" : `登入失敗：${e.message}`; }
    finally { login.disabled = false; }
  });
  logout.addEventListener("click", async () => {
    try { await signOut(auth); } catch (e) { error.textContent = `登出失敗：${e.message}`; }
  });
  onAuthStateChanged(auth, showUser, () => showUser(null));
}
