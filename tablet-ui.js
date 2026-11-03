// ============================================================
// Интерфейс планшета: переключатель «Авто / Палец и стилус / Мышь и клавиатура» и масштаб.
//  • Мышь — вид как на ноутбуке, обычное перетаскивание мышью.
//  • Палец/стилус — крупные кнопки и поля, перетаскивание долгим нажатием (≈0,3 с).
//  • Авто — если подключена мышь (any-pointer: fine), то режим «Мышь», иначе «Палец».
// Настройки хранятся только на этом устройстве (localStorage).
// ============================================================
(function () {
  const MODE_KEY = "lc_input_mode"; // auto | touch | mouse
  const SCALE_KEY = "lc_ui_scale"; // 100 | 115 | 130
  const get = (k, d) => {
    try {
      return localStorage.getItem(k) || d;
    } catch (e) {
      return d;
    }
  };
  const set = (k, v) => {
    try {
      localStorage.setItem(k, v);
    } catch (e) {}
  };
  let mode = get(MODE_KEY, "auto");
  let scale = parseInt(get(SCALE_KEY, "115"), 10) || 115;
  const mqFine = window.matchMedia("(any-pointer: fine)");
  const effective = () => (mode === "touch" ? "touch" : mode === "mouse" ? "mouse" : mqFine.matches ? "mouse" : "touch");

  const css = document.createElement("style");
  css.textContent = `
    html[data-input="touch"] button { min-height: 44px; }
    html[data-input="touch"] input:not([type=checkbox]):not([type=radio]):not([type=range]),
    html[data-input="touch"] select, html[data-input="touch"] textarea { min-height: 42px; font-size: 16px !important; }
    html[data-input="touch"] input[type=checkbox], html[data-input="touch"] input[type=radio] { width: 24px; height: 24px; }
    html[data-input="touch"] .lc-iconbtn { width: 40px; height: 40px; font-size: 17px; }
    html[data-input="touch"] .lc-row .lc-reveal { opacity: 1 !important; pointer-events: auto !important; }
    html[data-input="touch"] td, html[data-input="touch"] th { padding-top: 7px !important; padding-bottom: 7px !important; }
    html[data-input="touch"] [draggable="true"] { -webkit-user-select: none; user-select: none; -webkit-touch-callout: none; -webkit-user-drag: none; }
    html[data-input="touch"] [contenteditable] { -webkit-user-select: text; user-select: text; }
    #root [style*="flex-shrink: 0"][style*="flex-wrap: wrap"] { flex-shrink: 1 !important; max-width: 100%; }
    html { -webkit-tap-highlight-color: rgba(19,44,83,0.12); }
    body { overscroll-behavior-y: contain; }
    .top-row { flex-wrap: wrap; row-gap: 6px; }
    html[data-input="touch"] .topnav button { min-height: 44px; padding-left: 14px; padding-right: 14px; }
    html[data-input="touch"] .chip { min-height: 36px; }
    #lc-input-toggle { position: fixed; right: 10px; bottom: 10px; z-index: 5000; font-family: sans-serif; }
    #lc-input-toggle > button { border: 1px solid #b9cfe0; background: rgba(255,255,255,0.94); color: #132c53; border-radius: 22px; padding: 8px 14px; font-size: 13px; font-weight: 700; box-shadow: 0 2px 10px rgba(19,44,83,0.25); min-height: 40px; }
    #lc-input-panel { position: absolute; right: 0; bottom: 48px; width: 250px; background: #fff; color: #132c53; border: 1px solid #b9cfe0; border-radius: 14px; padding: 12px; box-shadow: 0 6px 24px rgba(19,44,83,0.3); display: none; font-size: 14px; }
    #lc-input-panel.open { display: block; }
    #lc-input-panel .lc-opt { display: block; width: 100%; text-align: left; margin: 0 0 6px; border: 1.5px solid #d7e6f0; background: #fff; border-radius: 10px; padding: 10px 12px; font-size: 14px; min-height: 44px; }
    #lc-input-panel .lc-opt.on { border-color: #132c53; background: #e7ecf5; font-weight: 700; }
    #lc-input-panel .lc-scale { display: flex; gap: 6px; margin-top: 4px; }
    #lc-input-panel .lc-scale button { flex: 1; border: 1.5px solid #d7e6f0; background: #fff; border-radius: 10px; min-height: 44px; font-size: 14px; }
    #lc-input-panel .lc-scale button.on { border-color: #132c53; background: #e7ecf5; font-weight: 700; }
    #lc-drag-ghost { position: fixed; z-index: 6000; pointer-events: none; opacity: 0.85; box-shadow: 0 8px 24px rgba(19,44,83,0.35); border-radius: 8px; transform: scale(1.03); }
    @media print { #lc-input-toggle, #lc-drag-ghost { display: none !important; } }
  `;
  document.head.appendChild(css);

  function apply() {
    const eff = effective();
    document.documentElement.dataset.input = eff;
    const root = document.getElementById("root");
    if (root) root.style.zoom = eff === "touch" && scale !== 100 ? String(scale / 100) : "";
    const btn = document.querySelector("#lc-input-toggle > button");
    if (btn) btn.textContent = (eff === "touch" ? "✋ " : "🖱 ") + (mode === "auto" ? "Авто" : eff === "touch" ? "Палец" : "Мышь");
    document.querySelectorAll("#lc-input-panel .lc-opt").forEach((el) => el.classList.toggle("on", el.dataset.mode === mode));
    document.querySelectorAll("#lc-input-panel .lc-scale button").forEach((el) => el.classList.toggle("on", parseInt(el.dataset.scale, 10) === scale));
  }
  mqFine.addEventListener && mqFine.addEventListener("change", apply);

  function buildToggle() {
    const wrap = document.createElement("div");
    wrap.id = "lc-input-toggle";
    wrap.innerHTML = `
      <div id="lc-input-panel">
        <div style="font-weight:700;margin-bottom:8px;">Управление</div>
        <button class="lc-opt" data-mode="auto">Авто (по подключённой мыши)</button>
        <button class="lc-opt" data-mode="touch">✋ Палец и стилус</button>
        <button class="lc-opt" data-mode="mouse">🖱 Мышь и клавиатура</button>
        <div style="font-weight:700;margin:8px 0 4px;">Размер (для режима «Палец»)</div>
        <div class="lc-scale"><button data-scale="100">100%</button><button data-scale="115">115%</button><button data-scale="130">130%</button></div>
      </div>
      <button type="button"></button>`;
    document.body.appendChild(wrap);
    const panel = wrap.querySelector("#lc-input-panel");
    wrap.querySelector(":scope > button").addEventListener("click", () => panel.classList.toggle("open"));
    panel.addEventListener("click", (e) => {
      const m = e.target.closest("[data-mode]");
      const s = e.target.closest("[data-scale]");
      if (m) {
        mode = m.dataset.mode;
        set(MODE_KEY, mode);
      }
      if (s) {
        scale = parseInt(s.dataset.scale, 10);
        set(SCALE_KEY, String(scale));
      }
      apply();
    });
    document.addEventListener("pointerdown", (e) => {
      if (!wrap.contains(e.target)) panel.classList.remove("open");
    });
    apply();
  }

  // ---------- перетаскивание пальцем (долгое нажатие на элемент с draggable) ----------
  // Chrome на Android не запускает перетаскивание HTML5 по касанию, поэтому повторяем его событиями:
  // dragstart → dragenter/dragover/dragleave → drop → dragend с общим DataTransfer.
  const HOLD_MS = 320;
  let pending = null; // { el, x, y, timer }
  let drag = null; // { el, dt, ghost, over, accepted, scrollTimer, x, y }

  const isTouchMode = () => document.documentElement.dataset.input === "touch";
  const fire = (target, type, dt, x, y) => {
    const ev = new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer: dt, clientX: x, clientY: y });
    return !target.dispatchEvent(ev); // true — обработчик разрешил сброс (вызвал preventDefault)
  };
  const underFinger = (x, y) => {
    const el = document.elementFromPoint(x, y);
    return el || document.body;
  };

  function startDrag(el, x, y) {
    const dt = new DataTransfer();
    dt.effectAllowed = "all";
    fire(el, "dragstart", dt, x, y);
    const rect = el.getBoundingClientRect();
    const ghost = el.cloneNode(true);
    ghost.id = "lc-drag-ghost";
    ghost.style.width = rect.width + "px";
    ghost.style.left = x - rect.width / 2 + "px";
    ghost.style.top = y - rect.height / 2 + "px";
    document.body.appendChild(ghost);
    drag = { el, dt, ghost, over: null, accepted: false, x, y, w: rect.width, h: rect.height, scrollTimer: null };
    drag.scrollTimer = setInterval(() => {
      if (!drag) return;
      const edge = 70;
      if (drag.y < edge) window.scrollBy(0, -16);
      else if (drag.y > window.innerHeight - edge) window.scrollBy(0, 16);
    }, 30);
    if (navigator.vibrate) navigator.vibrate(25);
  }
  function moveDrag(x, y) {
    drag.x = x;
    drag.y = y;
    drag.ghost.style.left = x - drag.w / 2 + "px";
    drag.ghost.style.top = y - drag.h / 2 + "px";
    const target = underFinger(x, y);
    if (target !== drag.over) {
      if (drag.over) fire(drag.over, "dragleave", drag.dt, x, y);
      fire(target, "dragenter", drag.dt, x, y);
      drag.over = target;
    }
    drag.accepted = fire(target, "dragover", drag.dt, x, y);
  }
  function endDrag(x, y, drop) {
    const d = drag;
    drag = null;
    clearInterval(d.scrollTimer);
    d.ghost.remove();
    if (drop && d.accepted) {
      const target = underFinger(x, y);
      fire(target, "drop", d.dt, x, y);
    } else if (d.over) {
      fire(d.over, "dragleave", d.dt, x, y);
    }
    fire(d.el, "dragend", d.dt, x, y);
  }
  function cancelPending() {
    if (pending) {
      clearTimeout(pending.timer);
      pending = null;
    }
  }

  document.addEventListener(
    "touchstart",
    (e) => {
      if (!isTouchMode() || e.touches.length !== 1 || drag) return;
      const el = e.target.closest && e.target.closest('[draggable="true"]');
      if (!el) return;
      const t = e.touches[0];
      cancelPending();
      pending = {
        el,
        x: t.clientX,
        y: t.clientY,
        timer: setTimeout(() => {
          const p = pending;
          pending = null;
          if (p) startDrag(p.el, p.x, p.y);
        }, HOLD_MS),
      };
    },
    { passive: true, capture: true }
  );
  document.addEventListener(
    "touchmove",
    (e) => {
      const t = e.touches[0];
      if (drag) {
        e.preventDefault();
        moveDrag(t.clientX, t.clientY);
      } else if (pending && Math.hypot(t.clientX - pending.x, t.clientY - pending.y) > 10) {
        cancelPending(); // палец поехал сразу — это прокрутка, а не перетаскивание
      }
    },
    { passive: false, capture: true }
  );
  const onEnd = (drop) => (e) => {
    cancelPending();
    if (drag) {
      e.preventDefault();
      const t = (e.changedTouches && e.changedTouches[0]) || { clientX: drag.x, clientY: drag.y };
      endDrag(t.clientX, t.clientY, drop);
    }
  };
  document.addEventListener("touchend", onEnd(true), { passive: false, capture: true });
  document.addEventListener("touchcancel", onEnd(false), { passive: false, capture: true });
  // собственное нативное перетаскивание/меню по долгому нажатию в режиме «Палец» мешает — гасим
  document.addEventListener(
    "dragstart",
    (e) => {
      if (isTouchMode() && e.isTrusted) {
        e.preventDefault();
        e.stopImmediatePropagation();
      }
    },
    true
  );
  document.addEventListener(
    "contextmenu",
    (e) => {
      if (isTouchMode() && (pending || drag || (e.target.closest && e.target.closest('[draggable="true"]')))) e.preventDefault();
    },
    true
  );

  const init = () => {
    buildToggle();
    // #root появляется сразу, но на всякий случай применяем масштаб ещё раз после старта приложения
    setTimeout(apply, 300);
  };
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
  window.lcTablet = { apply, getMode: () => mode, getEffective: effective };
})();
