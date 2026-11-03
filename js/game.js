/**
 * Strawberry Bane — faithful browser port of the Excel VBA game.
 * Gameplay rules mirror the VBA modules; timing approximates Application.OnTime / Wait.
 */
(() => {
  "use strict";

  // --- Excel-like layout (points ≈ CSS px) ---
  const COLS = 11; // B..L
  const COL_W = 70;
  const ROW1_H = 10;
  const ROW_H = 100;
  const STAGE_W = COL_W * COLS; // 770
  const STAGE_H = ROW1_H + ROW_H * 6; // 610
  const PLAYER_ROW = 6; // 0-based index among play rows; Excel row 7
  const ITEM_SIZE = 48;
  // Excel Bane shape ≈ 84×130 pt; keep slightly shorter than ROW_H so feet sit in row 7
  const BANE_W = 80;
  const BANE_H = 124;
  // VBA Wait(0.1s) effectively no-ops → DoEvents-rate loop. Base step was 1 pt/tick;
  // UX request: faster falls — keep ratio (normal vs increased 1 : 1.5), raise base.
  const MOVE_DELAY_MS = 12;
  const FALL_STEP_NORMAL = 3.5;
  const FALL_STEP_FAST = 5.25; // 3.5 * 1.5 — same VBA increased-step ratio
  const DROP_GAP_MS = 500;
  /** Dual-item burst: random 10s or 12s gap; second item after a short delay */
  const DUAL_INTERVAL_CHOICES_MS = [10000, 12000];
  const DUAL_SECOND_DELAY_MS = 400;

  function nextDualIntervalMs() {
    return DUAL_INTERVAL_CHOICES_MS[
      Math.floor(Math.random() * DUAL_INTERVAL_CHOICES_MS.length)
    ];
  }
  // HUD clock refreshes every 1s (UX). Game-end (300s) and speed boost (240s)
  // still keyed off elapsed time, same as VBA — only the display cadence changed.
  const TIMER_DISPLAY_MS = 1000;
  const GAME_LIMIT_S = 300;
  const SPEED_AT_S = 240;
  const SPEED_BOOST_MS = 30000;
  const FREEZE_MS = 3000;
  /** Jump trial: two Excel fields (rows) up, then sink back to catch row 7. */
  const JUMP_HEIGHT = 2 * ROW_H; // 200 px = two play rows (Sebastian trial)
  const JUMP_UP_MS = 200; // slightly longer for 2× height; still snappy
  const JUMP_DOWN_MS = 280;
  const JUMP_TICK_MS = 16;
  /** Best score (HUD). Full top-10 list lives under HS_LIST_KEY in localStorage. */
  const HS_KEY = "strawberryBaneHighscore";
  const HS_LIST_KEY = "strawberryBaneHighscoreList";
  const HS_MAX = 10;

  const ITEM_WEIGHTS = [
    ["Strawberry", 12],
    ["Banana", 10],
    ["Pork", 10],
    ["Bone", 10],
    ["Pizza", 10],
    ["Chicken", 8],
    ["Donut", 8],
    ["Ice", 8],
    ["Food", 4],
    ["Cat", 4],
    ["Football", 4],
    ["Bear", 3],
    ["Dinosaur", 3],
    ["Pineapple", 3],
    ["Gorilla", 3],
  ];

  const SPRITES = {
    Strawberry: "assets/sprites/strawberry.png",
    Banana: "assets/sprites/banana.png",
    Pork: "assets/sprites/pork.png",
    Bone: "assets/sprites/bone.png",
    Pizza: "assets/sprites/pizza.png",
    Chicken: "assets/sprites/chicken.png",
    Donut: "assets/sprites/donut.png",
    Ice: "assets/sprites/ice.png",
    Food: "assets/sprites/food.png",
    Cat: "assets/sprites/cat.png",
    Football: "assets/sprites/football.png",
    Bear: "assets/sprites/bear.png",
    Dinosaur: "assets/sprites/dinosaur.png",
    Pineapple: "assets/sprites/pineapple.png",
    Gorilla: "assets/sprites/gorilla.png",
    Bane: "assets/sprites/bane.png",
    Poodle: "assets/sprites/poodle.png",
  };

  const SOUND_FILES = {
    eat_crunch: "eat_crunch.wav",
    bark_dislike: "bark_dislike.wav",
    meal_burp: "meal_burp.wav",
    cat_sound: "cat_sound.wav",
    hum_football: "hum_football.wav",
    bear_noise: "bear_noise.wav",
    dinosaur_noise: "dinosaur_noise.wav",
    gorilla_entry: "gorilla_entry.wav",
  };

  // Sound search paths (game folder + store media)
  const SOUND_BASES = [
    "assets/sounds/",
    "../../media/sounds/",
    "/cursor/stores/self/media/sounds/",
  ];

  // --- DOM ---
  const stage = document.getElementById("stage");
  const stageScaleEl = document.getElementById("stage-scale");
  const stageWrap = document.getElementById("stage-wrap");
  const entitiesEl = document.getElementById("entities");
  const livesEl = document.getElementById("lives");
  const scoreEl = document.getElementById("score");
  const timeEl = document.getElementById("time");
  const highscoreEl = document.getElementById("highscore");
  const hud = document.getElementById("hud");
  const statusEl = document.getElementById("status");
  const overlayGoodLuck = document.getElementById("overlay-goodluck");
  const overlayGameOver = document.getElementById("overlay-gameover");
  const freezeBanner = document.getElementById("freeze-banner");
  const speedBanner = document.getElementById("speed-banner");
  const modal = document.getElementById("modal");
  const modalImg = document.getElementById("modal-img");
  const modalActions = document.getElementById("modal-actions");
  const toast = document.getElementById("toast");

  const btnStart = document.getElementById("btn-start");
  const btnIntro = document.getElementById("btn-intro");
  const btnHelp = document.getElementById("btn-help");
  const btnCredits = document.getElementById("btn-credits");
  const btnHighscore = document.getElementById("btn-highscore");
  const hsModal = document.getElementById("hs-modal");
  const hsListEl = document.getElementById("hs-list");
  const nameModal = document.getElementById("name-modal");
  const nameForm = document.getElementById("name-form");
  const nameInput = document.getElementById("player-name");
  const nameScoreEl = document.getElementById("name-score");

  // Sync CSS vars (logical size — never make these responsive percentages)
  document.documentElement.style.setProperty("--stage-w", STAGE_W + "px");
  document.documentElement.style.setProperty("--stage-h", STAGE_H + "px");
  document.documentElement.style.setProperty("--col-w", COL_W + "px");
  document.documentElement.style.setProperty("--row-h", ROW_H + "px");

  // Stage must be focusable so arrow keys reach the game after clicking Start
  stage.tabIndex = 0;
  stage.setAttribute("role", "application");
  stage.setAttribute("aria-label", "Strawberry Bane Spielfeld");

  /**
   * Stage stays a fixed 770×610 game grid (no sprite distortion).
   * The whole #app column is uniformly scaled to fit the viewport
   * (no page scroll on typical laptop/desktop windows).
   */
  const appEl = document.getElementById("app");
  const fitRoot = document.getElementById("fit-root");

  function fitStageNative() {
    if (!stageScaleEl) return;
    stage.style.transform = "scale(1)";
    stageScaleEl.style.width = STAGE_W + "px";
    stageScaleEl.style.height = STAGE_H + "px";
    document.documentElement.style.setProperty("--logo-w", STAGE_W + "px");
  }

  function fitViewport() {
    if (!appEl || !fitRoot) return;
    fitStageNative();
    // Reset before measuring unscaled layout size
    appEl.style.transform = "none";
    appEl.style.marginTop = "0";
    appEl.style.marginLeft = "0";

    const pad = 10;
    const availW = Math.max(200, window.innerWidth - pad * 2);
    const availH = Math.max(200, window.innerHeight - pad * 2);
    const w = appEl.offsetWidth;
    const h = appEl.offsetHeight;
    if (!w || !h) return;

    const s = Math.min(1, availW / w, availH / h);
    const scaledW = w * s;
    const scaledH = h * s;
    // Explicit viewport centering so logo/stage/footer share one axis
    const x = (window.innerWidth - scaledW) / 2;
    const y = Math.max(pad, (window.innerHeight - scaledH) / 2);
    appEl.style.transformOrigin = "top left";
    appEl.style.transform = "translate(" + x + "px, " + y + "px) scale(" + s + ")";
  }

  window.addEventListener("resize", fitViewport);

  /** Toggle visibility; pairs with CSS that forces [hidden] { display:none !important } */
  function setShown(el, shown) {
    if (!el) return;
    if (shown) {
      el.hidden = false;
      el.removeAttribute("hidden");
    } else {
      el.hidden = true;
      el.setAttribute("hidden", "");
    }
  }

  function focusStage() {
    try {
      stage.focus({ preventScroll: true });
    } catch (_) {
      stage.focus();
    }
  }

  // --- State (mirrors VBA globals) ---
  const state = {
    bGameRunning: false,
    bFormRunning: false,
    bGameEnded: false,
    bMoveDownRunning: false,
    keyboardLocked: false,
    unlockTime: 0,
    increasedStepActive: false,
    startTime: 0,
    lives: 3,
    score: 0,
    highscore: Number(localStorage.getItem(HS_KEY) || 0),
    baneCol: 5, // 0..10 = Excel cols B..L; G = index 5
    /** Pixels raised above ground (row 7) placement; 0 = standing. */
    baneJumpOffset: 0,
    baneJumping: false,
    jumpAnimId: null,
    fallingItems: [],
    poodleActive: false,
    poodle: null,
    timers: new Set(),
    moveLoopId: null,
    dropPending: false,
    moveLoopActive: false,
    lastDualAt: 0,
    nextDualInMs: 10000,
    pendingSecondary: false,
  };

  // --- Highscore list (localStorage JSON “database”) ---
  function loadHighscoreList() {
    try {
      const raw = localStorage.getItem(HS_LIST_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed
        .filter((e) => e && typeof e.score === "number")
        .map((e) => ({
          name: String(e.name || "Player").slice(0, 20),
          score: Math.max(0, Math.round(e.score)),
          date: e.date || "",
        }))
        .sort((a, b) => b.score - a.score)
        .slice(0, HS_MAX);
    } catch (_) {
      return [];
    }
  }

  function saveHighscoreList(list) {
    const cleaned = list
      .slice()
      .sort((a, b) => b.score - a.score)
      .slice(0, HS_MAX);
    localStorage.setItem(HS_LIST_KEY, JSON.stringify(cleaned));
    const best = cleaned.length ? cleaned[0].score : 0;
    state.highscore = best;
    localStorage.setItem(HS_KEY, String(best));
    highscoreEl.textContent = String(best);
    return cleaned;
  }

  /** Migrate legacy single highscore number into a list entry once. */
  function migrateLegacyHighscore() {
    const list = loadHighscoreList();
    if (list.length) return list;
    const legacy = Number(localStorage.getItem(HS_KEY) || 0);
    if (legacy > 0) {
      return saveHighscoreList([
        { name: "Player", score: legacy, date: new Date().toISOString().slice(0, 10) },
      ]);
    }
    return list;
  }

  function qualifiesForHighscoreList(score) {
    if (!(score > 0)) return false;
    const list = loadHighscoreList();
    if (list.length < HS_MAX) return true;
    return score > list[list.length - 1].score;
  }

  function addHighscoreEntry(name, score) {
    const list = loadHighscoreList();
    list.push({
      name: (name || "Player").trim().slice(0, 20) || "Player",
      score: Math.round(score),
      date: new Date().toISOString().slice(0, 10),
    });
    return saveHighscoreList(list);
  }

  function renderHighscoreList() {
    const list = loadHighscoreList();
    hsListEl.innerHTML = "";
    if (!list.length) {
      const empty = document.createElement("li");
      empty.className = "hs-empty";
      empty.textContent = "Noch keine Einträge — spiel eine Runde!";
      hsListEl.appendChild(empty);
      return;
    }
    list.forEach((entry, i) => {
      const li = document.createElement("li");
      li.innerHTML =
        `<span class="hs-rank">${i + 1}.</span>` +
        `<span class="hs-name"></span>` +
        `<span class="hs-meta"></span>`;
      li.querySelector(".hs-name").textContent = entry.name;
      li.querySelector(".hs-meta").textContent =
        entry.score + (entry.date ? " · " + entry.date : "");
      hsListEl.appendChild(li);
    });
  }

  function showHighscoreModal() {
    renderHighscoreList();
    setShown(hsModal, true);
  }

  function hideHighscoreModal() {
    setShown(hsModal, false);
  }

  let pendingHsScore = 0;

  function showNameModal(score) {
    pendingHsScore = score;
    nameScoreEl.textContent = String(score);
    nameInput.value = "";
    setShown(nameModal, true);
    schedule(() => nameInput.focus(), 50);
  }

  function hideNameModal() {
    setShown(nameModal, false);
    pendingHsScore = 0;
  }

  function finishNameEntry(name) {
    const score = pendingHsScore;
    hideNameModal();
    if (score > 0) {
      addHighscoreEntry(name, score);
      showToast("Highscore gespeichert: " + (name || "Player") + " — " + score);
    }
  }

  // Seed HUD best from list / legacy single value
  {
    const seeded = migrateLegacyHighscore();
    if (seeded.length) state.highscore = seeded[0].score;
  }
  highscoreEl.textContent = String(state.highscore);

  // --- Helpers ---
  function colLeft(colIndex) {
    return colIndex * COL_W;
  }

  function rowTop(excelRow /* 1..7 */) {
    if (excelRow <= 1) return 0;
    return ROW1_H + (excelRow - 2) * ROW_H;
  }

  function centerInCell(colIndex, excelRow, w, h) {
    const left = colLeft(colIndex) + (COL_W - w) / 2;
    let top;
    if (excelRow === 7) {
      // Excel Bane is taller than the row — feet on row 7, body extends into row 6
      const row7 = rowTop(7);
      top = row7 + ROW_H - h; // bottom-align in catch row
      if (top < rowTop(6)) top = rowTop(6); // keep mostly in view
    } else if (excelRow === 1) {
      top = rowTop(1) + Math.max(0, (ROW1_H - h) / 2);
    } else {
      top = rowTop(excelRow) + Math.max(0, (ROW_H - h) / 2);
    }
    return { left, top };
  }

  function schedule(fn, ms) {
    const id = setTimeout(() => {
      state.timers.delete(id);
      fn();
    }, ms);
    state.timers.add(id);
    return id;
  }

  function clearAllTimers() {
    for (const id of state.timers) clearTimeout(id);
    state.timers.clear();
    if (state.moveLoopId) {
      clearTimeout(state.moveLoopId);
      state.moveLoopId = null;
    }
    state.jumpAnimId = null;
  }

  function setStatus(msg) {
    statusEl.textContent = msg || "";
  }

  function showToast(msg, ms = 2500) {
    toast.textContent = msg;
    setShown(toast, true);
    schedule(() => {
      setShown(toast, false);
    }, ms);
  }

  function updateHud() {
    livesEl.textContent = String(state.lives);
    scoreEl.textContent = String(state.score);
    setShown(
      freezeBanner,
      state.keyboardLocked && performance.now() < state.unlockTime
    );
    setShown(speedBanner, state.increasedStepActive);
  }

  function formatTime(elapsedS) {
    const m = Math.floor(elapsedS / 60);
    const s = Math.floor(elapsedS % 60);
    return String(m).padStart(2, "0") + ":" + String(s).padStart(2, "0");
  }

  function getCurrentStep() {
    return state.increasedStepActive ? FALL_STEP_FAST : FALL_STEP_NORMAL;
  }

  function weightedRandomItem() {
    let total = 0;
    for (const [, w] of ITEM_WEIGHTS) total += w;
    let r = Math.floor(Math.random() * total) + 1;
    let cum = 0;
    for (const [name, w] of ITEM_WEIGHTS) {
      cum += w;
      if (r <= cum) return name;
    }
    return ITEM_WEIGHTS[0][0];
  }

  function getRandomDirection() {
    return Math.floor(Math.random() * 3) - 1; // -1,0,1
  }

  // --- Sound ---
  const audioCache = new Map();

  function playSound(fileName) {
    // Try each base; fail silently like VBA Dir check
    const tryPlay = (bases, i) => {
      if (i >= bases.length) return;
      const url = bases[i] + fileName;
      let audio = audioCache.get(url);
      if (!audio) {
        audio = new Audio(url);
        audioCache.set(url, audio);
        audio.addEventListener("error", () => {
          audioCache.delete(url);
          tryPlay(bases, i + 1);
        }, { once: true });
      }
      audio.currentTime = 0;
      audio.play().catch(() => tryPlay(bases, i + 1));
    };
    tryPlay(SOUND_BASES, 0);
  }

  // --- Entities ---
  function makeEl(className, spriteKey) {
    const el = document.createElement("div");
    el.className = "entity " + className;
    el.style.backgroundImage = `url("${SPRITES[spriteKey]}")`;
    entitiesEl.appendChild(el);
    return el;
  }

  let baneEl = null;

  function setupBane() {
    if (!baneEl) {
      baneEl = document.createElement("div");
      baneEl.className = "entity bane";
      // <img> is more reliable than background-image across file:// / viewers
      const img = document.createElement("img");
      img.src = SPRITES.Bane;
      img.alt = "Bane";
      img.draggable = false;
      baneEl.appendChild(img);
      entitiesEl.appendChild(baneEl);
    }
    state.baneCol = 5;
    resetJump();
    placeBane();
  }

  function placeBane() {
    if (!baneEl) return;
    const { left, top } = centerInCell(state.baneCol, 7, BANE_W, BANE_H);
    baneEl.style.left = left + "px";
    baneEl.style.top = top - state.baneJumpOffset + "px";
    baneEl.style.width = BANE_W + "px";
    baneEl.style.height = BANE_H + "px";
    baneEl.style.zIndex = "8";
  }

  function clearJumpAnim() {
    if (state.jumpAnimId) {
      clearTimeout(state.jumpAnimId);
      state.timers.delete(state.jumpAnimId);
      state.jumpAnimId = null;
    }
  }

  function resetJump() {
    clearJumpAnim();
    state.baneJumping = false;
    state.baneJumpOffset = 0;
  }

  /**
   * Animate baneJumpOffset from `from` to `to` over `durationMs`, then call `onDone`.
   * placeBane each tick so left/right mid-jump and baneBox() stay in sync.
   */
  function animateJumpOffset(from, to, durationMs, onDone) {
    const started = performance.now();
    const tick = () => {
      if (!state.bFormRunning) {
        resetJump();
        placeBane();
        return;
      }
      const t = Math.min(1, (performance.now() - started) / durationMs);
      // Ease-out on ascent, ease-in on descent (sink)
      const eased = to > from ? 1 - Math.pow(1 - t, 2) : t * t;
      state.baneJumpOffset = from + (to - from) * eased;
      placeBane();
      if (t < 1) {
        state.jumpAnimId = schedule(tick, JUMP_TICK_MS);
      } else {
        state.jumpAnimId = null;
        state.baneJumpOffset = to;
        placeBane();
        if (onDone) onDone();
      }
    };
    tick();
  }

  /** Two-row hop (trial): rise JUMP_HEIGHT, then sink back to the catch row. */
  function startJump() {
    if (!state.bFormRunning) return;
    const now = performance.now();
    if (now < state.unlockTime) {
      state.keyboardLocked = true;
      return;
    }
    if (state.keyboardLocked) {
      state.keyboardLocked = false;
      updateHud();
    }
    if (state.baneJumping) return;

    state.baneJumping = true;
    clearJumpAnim();
    animateJumpOffset(0, JUMP_HEIGHT, JUMP_UP_MS, () => {
      animateJumpOffset(JUMP_HEIGHT, 0, JUMP_DOWN_MS, () => {
        state.baneJumping = false;
        state.baneJumpOffset = 0;
        placeBane();
      });
    });
    updateHud();
  }

  /** Bane only while the match is running (after Good Luck, before Game Over). */
  function setBaneVisible(visible) {
    if (!baneEl) return;
    baneEl.style.visibility = visible ? "visible" : "hidden";
    baneEl.style.opacity = visible ? "1" : "0";
    baneEl.setAttribute("aria-hidden", visible ? "false" : "true");
  }

  function baneBox() {
    return {
      left: parseFloat(baneEl.style.left) || 0,
      top: parseFloat(baneEl.style.top) || 0,
      width: BANE_W,
      height: BANE_H,
    };
  }

  /** True while rising/sinking — no column change mid-air (straight-up jump only). */
  function isBaneAirborne() {
    return state.baneJumping || state.baneJumpOffset > 0;
  }

  function moveImage(direction) {
    if (!state.bFormRunning) return;
    // Sebastian: left/right only on the catch row — no air strafe / side-jump.
    if (isBaneAirborne()) return;
    const now = performance.now();
    // Unlock only when the latest unlockTime has elapsed (every Pineapple
    // catch restarts unlockTime — never honor a stale earlier window).
    if (now < state.unlockTime) {
      state.keyboardLocked = true;
      return;
    }
    if (state.keyboardLocked) {
      state.keyboardLocked = false;
      updateHud();
    }

    if (direction === "Left" && state.baneCol > 0) {
      state.baneCol -= 1;
      placeBane();
    } else if (direction === "Right" && state.baneCol < COLS - 1) {
      state.baneCol += 1;
      placeBane();
    }
    updateHud();
  }

  function checkCollisionWithBane(item) {
    const b = baneBox();
    // VBA: shp.Top + shp.Height > bane.Top AND horizontal overlap (no full vertical AABB)
    return (
      item.top + item.height > b.top &&
      item.left + item.width > b.left &&
      item.left < b.left + b.width
    );
  }

  function clearFallingShapes() {
    for (const item of state.fallingItems) {
      try {
        item.el.remove();
      } catch (_) {}
    }
    state.fallingItems = [];
    state.pendingSecondary = false;
    entitiesEl.querySelectorAll(".entity.item").forEach((n) => n.remove());
  }

  // --- Collision effects ---
  function updateCellValues(itemName, collisionOccurred) {
    if (collisionOccurred) {
      switch (itemName) {
        case "Strawberry":
        case "Banana":
        case "Pork":
        case "Bone": {
          playSound(SOUND_FILES.eat_crunch);
          const add =
            itemName === "Strawberry" ? 25 :
            itemName === "Banana" ? 15 :
            itemName === "Pork" ? 10 : 5;
          state.score += add;
          break;
        }
        case "Pizza":
        case "Chicken":
        case "Donut":
        case "Ice":
          playSound(SOUND_FILES.bark_dislike);
          state.lives -= 1;
          break;
        case "Food":
          playSound(SOUND_FILES.meal_burp);
          state.lives += 1;
          break;
        case "Cat":
          playSound(SOUND_FILES.cat_sound);
          state.lives = 0;
          break;
        case "Football":
          playSound(SOUND_FILES.hum_football);
          moveBaneRandom();
          state.lives -= 1;
          state.score += 100;
          break;
        case "Bear":
          playSound(SOUND_FILES.bear_noise);
          state.score = Math.round(state.score * 2);
          break;
        case "Dinosaur":
          playSound(SOUND_FILES.dinosaur_noise);
          state.score = Math.round(state.score / 2);
          break;
        case "Gorilla":
          playSound(SOUND_FILES.gorilla_entry);
          state.score += 1000;
          activateIncreasedStep();
          break;
        case "Pineapple":
          playSound(SOUND_FILES.eat_crunch);
          state.lives += 1;
          // Every catch restarts a full FREEZE_MS lock from *this* moment
          // (cancel any prior unlock timeout so it cannot clear the new freeze early).
          activateKeyboardFreeze();
          break;
        default:
          break;
      }
    } else if (
      itemName === "Strawberry" ||
      itemName === "Banana" ||
      itemName === "Pork" ||
      itemName === "Bone"
    ) {
      // Missed loved item (life loss) — same bark as hated catch
      playSound(SOUND_FILES.bark_dislike);
      state.lives -= 1;
    }

    updateHud();

    if (state.lives === 0 || state.score < 0) {
      endGame();
    }
  }

  function moveBaneRandom() {
    // VBA MoveBaneRandom "B7","L7"
    if (state.baneCol === 0) state.baneCol = COLS - 1;
    else if (state.baneCol === COLS - 1) state.baneCol = 0;
    else state.baneCol = Math.random() < 0.5 ? 0 : COLS - 1;
    placeBane();
  }

  // --- Speed boost ---
  let speedDeactivateId = null;
  // --- Pineapple keyboard freeze (mirrors speed-boost cancel/reschedule) ---
  let freezeUnlockId = null;

  function clearFreezeUnlockTimeout() {
    if (freezeUnlockId) {
      clearTimeout(freezeUnlockId);
      state.timers.delete(freezeUnlockId);
      freezeUnlockId = null;
    }
  }

  /**
   * Restart the full FREEZE_MS keyboard lock from now.
   * Must cancel any prior unlock timeout — otherwise catch #1's timer can fire
   * during catch #2's window and clear keyboardLocked early (feels like the
   * second Pineapple cancelled the freeze).
   */
  function activateKeyboardFreeze() {
    const now = performance.now();
    state.keyboardLocked = true;
    state.unlockTime = now + FREEZE_MS;
    clearFreezeUnlockTimeout();
    freezeUnlockId = schedule(() => {
      freezeUnlockId = null;
      // Only unlock if this catch's window has actually elapsed (defensive).
      if (performance.now() >= state.unlockTime) {
        state.keyboardLocked = false;
      }
      updateHud();
    }, FREEZE_MS);
    updateHud();
  }

  function activateIncreasedStep() {
    state.increasedStepActive = true;
    if (speedDeactivateId) {
      clearTimeout(speedDeactivateId);
      state.timers.delete(speedDeactivateId);
    }
    speedDeactivateId = schedule(() => {
      state.increasedStepActive = false;
      speedDeactivateId = null;
      updateHud();
    }, SPEED_BOOST_MS);
    updateHud();
  }

  // --- Drop / MoveDown (supports occasional dual falling items) ---
  function isSpawnColEmpty(colIndex) {
    const cellLeft = colLeft(colIndex);
    const cellRight = cellLeft + COL_W;
    const cellTop = 0;
    const cellBottom = ROW1_H + ITEM_SIZE;
    for (const f of state.fallingItems) {
      if (
        f.top < cellBottom &&
        f.top + f.height > cellTop &&
        f.left + f.width > cellLeft &&
        f.left < cellRight
      ) {
        return false;
      }
    }
    return true;
  }

  function spawnFallingItem() {
    const itemName = weightedRandomItem();
    // Prefer an empty column; try a few random picks
    let col = Math.floor(Math.random() * 11);
    let found = isSpawnColEmpty(col);
    if (!found) {
      for (let i = 0; i < 11; i++) {
        const c = (col + i) % 11;
        if (isSpawnColEmpty(c)) {
          col = c;
          found = true;
          break;
        }
      }
    }
    if (!found) return false;

    const el = makeEl("item", itemName);
    const pos = centerInCell(col, 1, ITEM_SIZE, ITEM_SIZE);
    const top = Math.max(0, pos.top);
    el.style.left = pos.left + "px";
    el.style.top = top + "px";
    el.style.width = ITEM_SIZE + "px";
    el.style.height = ITEM_SIZE + "px";

    // lastDirection -1/0/1 → zigzag or straight (VBA GetRandomDirection)
    state.fallingItems.push({
      name: itemName,
      el,
      left: pos.left,
      top,
      width: ITEM_SIZE,
      height: ITEM_SIZE,
      lastDirection: getRandomDirection(),
    });
    return true;
  }

  function ensureMoveLoop() {
    if (state.moveLoopActive) return;
    state.moveLoopActive = true;
    state.bMoveDownRunning = true;
    moveDownTick();
  }

  function dropItems() {
    if (!state.bFormRunning) return;
    // Primary cycle waits until the field is clear (and no dual second pending)
    if (state.fallingItems.length > 0 || state.pendingSecondary) {
      scheduleNextDrop();
      return;
    }

    if (!spawnFallingItem()) {
      scheduleNextDrop();
      return;
    }
    ensureMoveLoop();

    const now = performance.now();
    if (now - state.lastDualAt >= state.nextDualInMs) {
      state.lastDualAt = now;
      state.nextDualInMs = nextDualIntervalMs(); // 10s or 12s until next opportunity
      state.pendingSecondary = true;
      schedule(() => {
        state.pendingSecondary = false;
        if (!state.bFormRunning) return;
        spawnFallingItem();
        ensureMoveLoop();
      }, DUAL_SECOND_DELAY_MS);
    }
  }

  function scheduleNextDrop() {
    if (!state.bFormRunning) return;
    if (state.dropPending) return;
    state.dropPending = true;
    schedule(() => {
      state.dropPending = false;
      dropItems();
    }, DROP_GAP_MS);
  }

  /** Advance one falling item one step. Returns true if still active. */
  function stepFallingItem(item) {
    const step = getCurrentStep();
    let dir = state.increasedStepActive ? 0 : item.lastDirection;
    let newLeft = item.left + dir * step * 2;
    const leftBound = 0;
    const rightBound = STAGE_W;
    if (newLeft < leftBound) {
      newLeft = leftBound;
      item.lastDirection = 1;
    } else if (newLeft + item.width > rightBound) {
      newLeft = rightBound - item.width;
      item.lastDirection = -1;
    }

    item.left = newLeft;
    item.top += step;
    item.el.style.left = item.left + "px";
    item.el.style.top = item.top + "px";

    const stopLine = rowTop(7);
    const approxTopLeftRow = itemTopLeftExcelRow(item);

    // While airborne, apply the same AABB catch as on the ground using the
    // current jumped baneBox() — items can meet Bane one row higher.
    if (state.baneJumpOffset > 0 && checkCollisionWithBane(item)) {
      item.el.remove();
      updateCellValues(item.name, true);
      return false;
    }

    if (approxTopLeftRow === 6) {
      if (checkCollisionWithBane(item)) {
        item.el.remove();
        updateCellValues(item.name, true);
        return false;
      }
      const row6Top = rowTop(6);
      if (item.top + item.height > row6Top + ROW_H) {
        item.el.remove();
        updateCellValues(item.name, false);
        return false;
      }
    }

    if (item.top + item.height > stopLine + ROW_H) {
      const hit = checkCollisionWithBane(item);
      item.el.remove();
      updateCellValues(item.name, hit);
      return false;
    }
    return true;
  }

  function moveDownTick() {
    if (!state.bFormRunning || !state.bMoveDownRunning) {
      state.moveLoopActive = false;
      return;
    }

    if (state.fallingItems.length === 0) {
      state.moveLoopActive = false;
      // Wait for dual secondary if still pending; else continue primary cycle
      if (state.pendingSecondary) {
        schedule(() => {
          if (!state.bFormRunning) return;
          if (state.fallingItems.length > 0) ensureMoveLoop();
          else if (!state.pendingSecondary) scheduleNextDrop();
        }, DUAL_SECOND_DELAY_MS + 50);
      } else {
        scheduleNextDrop();
      }
      return;
    }

    const still = [];
    for (const item of state.fallingItems) {
      if (stepFallingItem(item)) still.push(item);
      if (!state.bFormRunning) break; // EndGame may have fired mid-loop
    }
    state.fallingItems = still;

    if (!state.bFormRunning) {
      state.moveLoopActive = false;
      return;
    }

    state.moveLoopId = schedule(moveDownTick, MOVE_DELAY_MS);
  }

  function itemTopLeftExcelRow(item) {
    // Map Y to Excel row of TopLeftCell
    if (item.top < ROW1_H) return 1;
    const y = item.top - ROW1_H;
    const r = Math.floor(y / ROW_H) + 2; // row 2+
    return Math.min(7, Math.max(1, r));
  }

  // --- Timer ---
  function updateTimer() {
    if (!state.bGameRunning) return;
    const elapsed = (performance.now() - state.startTime) / 1000;
    if (elapsed >= GAME_LIMIT_S) {
      timeEl.textContent = "05:00";
      endGame();
      return;
    }
    timeEl.textContent = formatTime(elapsed);
    // Keep freeze banner in sync even when no keys/collisions fire updateHud.
    if (state.keyboardLocked && performance.now() >= state.unlockTime) {
      state.keyboardLocked = false;
    }
    updateHud();
    // 4-minute speed boost: still based on elapsed time (checked each display tick)
    if (elapsed >= SPEED_AT_S && !state.increasedStepActive) {
      activateIncreasedStep();
    }
    schedule(updateTimer, TIMER_DISPLAY_MS);
  }

  // --- Poodle ---
  function startPoodleMovement() {
    if (state.poodleActive) return;
    const initialDelay = 3;
    const nextDrop = Math.floor(Math.random() * 5) + 15 + initialDelay; // 18..23
    state.poodleActive = true;
    schedule(createAndMovePoodle, nextDrop * 1000);
  }

  function createAndMovePoodle() {
    if (!state.bFormRunning) {
      state.poodleActive = false;
      return;
    }
    const startRow = Math.floor(Math.random() * 5) + 2; // Excel rows 2..6
    const el = makeEl("poodle", "Poodle");
    let col = 0;
    const place = () => {
      // bottom-aligned in cell like VBA
      const cellTop = rowTop(startRow);
      const top = cellTop + ROW_H - ITEM_SIZE;
      const left = colLeft(col) + (COL_W - ITEM_SIZE) / 2;
      el.style.left = left + "px";
      el.style.top = top + "px";
      el.style.width = ITEM_SIZE + "px";
      el.style.height = ITEM_SIZE + "px";
    };
    place();
    state.poodle = { el, col, startRow };

    const stepPoodle = () => {
      if (!state.bFormRunning) {
        deletePoodle();
        return;
      }
      if (!state.poodle) return;
      if (state.poodle.col < COLS - 1) {
        state.poodle.col += 1;
        col = state.poodle.col;
        place();
        schedule(stepPoodle, 1000);
      } else {
        schedule(deletePoodle, 1000);
      }
    };
    // First position is col 0; after 1s move to next (VBA increments after placing)
    schedule(stepPoodle, 1000);
  }

  function deletePoodle() {
    if (state.poodle) {
      state.poodle.el.remove();
      state.poodle = null;
    }
    state.poodleActive = false;
    if (state.bFormRunning) startPoodleMovement();
  }

  // --- Game flow ---
  function setMenuEnabled(on) {
    btnStart.disabled = !on;
    btnIntro.disabled = false;
    btnHelp.disabled = false;
    btnCredits.disabled = false;
  }

  function showUserForm() {
    state.lives = 3;
    state.score = 0;
    state.bGameRunning = true;
    state.bGameEnded = false;
    state.keyboardLocked = false;
    state.unlockTime = 0;
    clearFreezeUnlockTimeout();
    state.increasedStepActive = false;
    state.startTime = performance.now();
    state.lastDualAt = performance.now();
    state.nextDualInMs = nextDualIntervalMs(); // first dual after random 10s or 12s
    state.pendingSecondary = false;
    state.moveLoopActive = false;
    state.dropPending = false;
    state.fallingItems = [];
    timeEl.textContent = "00:00";
    setShown(hud, true);
    updateHud();
    setShown(overlayGoodLuck, false);
    setShown(overlayGameOver, false);
    setupBane();
    setBaneVisible(true);
    state.bFormRunning = true;
    state.bMoveDownRunning = true;
    setStatus("Spiel läuft — Pfeiltasten ← → ↑ springen · Esc beendet");
    focusStage();
    updateTimer();
    dropItems();
    startPoodleMovement();
    // VBA resets lives to 3 again after StartPoodleMovement
    state.lives = 3;
    updateHud();
    focusStage();
    fitViewport();
  }

  function beginStartSequence() {
    if (state.bGameRunning || state.bFormRunning) return;
    setMenuEnabled(false);
    btnStart.disabled = true;
    clearAllTimers();
    clearFallingShapes();
    if (baneEl) baneEl.remove();
    baneEl = null;
    entitiesEl.querySelectorAll(".entity.poodle").forEach((n) => n.remove());
    state.poodle = null;
    state.poodleActive = false;
    setShown(hud, true);
    setupBane();
    setBaneVisible(false); // hidden until Good Luck finishes
    setShown(overlayGameOver, false);
    setShown(overlayGoodLuck, true);
    setStatus("Good Luck!");
    fitViewport();
    // ShowGoodLuck 2s → HideGoodLuck → ShowUserForm (then show Bane)
    schedule(() => {
      setShown(overlayGoodLuck, false);
      showUserForm();
    }, 2000);
  }

  function endGame() {
    if (state.bGameEnded && !state.bFormRunning && !state.bGameRunning) return;
    state.bFormRunning = false;
    state.bGameRunning = false;
    state.bGameEnded = true;
    state.bMoveDownRunning = false;
    state.moveLoopActive = false;
    state.pendingSecondary = false;
    state.increasedStepActive = false;
    resetJump();
    setBaneVisible(false);
    clearAllTimers();
    speedDeactivateId = null;
    freezeUnlockId = null;
    state.keyboardLocked = false;
    state.unlockTime = 0;

    // Brief delay then clear (VBA waits 1s)
    schedule(() => {
      clearFallingShapes();
      if (state.poodle) {
        state.poodle.el.remove();
        state.poodle = null;
      }
      state.poodleActive = false;
      setupBane();
      setBaneVisible(false);

      const finalScore = state.score;
      const askName = qualifiesForHighscoreList(finalScore);

      setShown(overlayGameOver, true);
      setStatus("Game Over");
      updateHud();

      schedule(() => {
        setShown(overlayGameOver, false);
        setBaneVisible(false);
        setMenuEnabled(true);
        btnStart.disabled = false;
        setStatus("Bereit — Start Game für eine neue Runde");
        if (askName) {
          showNameModal(finalScore);
        } else if (finalScore > state.highscore) {
          // Edge: beat HUD best but somehow not list — still refresh best display
          state.highscore = finalScore;
          localStorage.setItem(HS_KEY, String(finalScore));
          highscoreEl.textContent = String(finalScore);
        }
      }, 3000);
    }, 200);
  }

  // --- Modals (UserForms) ---
  function openModal(imageSrc, buttons) {
    modalImg.src = imageSrc;
    modalActions.innerHTML = "";
    for (const b of buttons) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = b.label;
      btn.addEventListener("click", () => {
        if (b.close) closeModal();
        if (b.action) b.action();
      });
      modalActions.appendChild(btn);
    }
    setShown(modal, true);
  }

  function closeModal() {
    setShown(modal, false);
    modalImg.removeAttribute("src");
    if (state.bFormRunning) focusStage();
  }

  function isModalOpen() {
    return !modal.hidden || !hsModal.hidden || !nameModal.hidden;
  }

  document.getElementById("modal-close").addEventListener("click", closeModal);
  modal.addEventListener("click", (e) => {
    if (e.target === modal) closeModal();
  });

  document.getElementById("hs-close").addEventListener("click", hideHighscoreModal);
  document.getElementById("hs-modal-x").addEventListener("click", hideHighscoreModal);
  hsModal.addEventListener("click", (e) => {
    if (e.target === hsModal) hideHighscoreModal();
  });

  nameForm.addEventListener("submit", (e) => {
    e.preventDefault();
    finishNameEntry(nameInput.value);
  });
  document.getElementById("name-skip").addEventListener("click", () => {
    finishNameEntry("Player");
  });
  document.getElementById("name-modal-x").addEventListener("click", () => {
    finishNameEntry("Player");
  });
  nameModal.addEventListener("click", (e) => {
    if (e.target === nameModal) finishNameEntry("Player");
  });

  function showIntro() {
    openModal("assets/forms/intro.jpg", [
      { label: "Learn More", action: () => { closeModal(); showLearnMore(); } },
      { label: "Let’s Play", close: true },
    ]);
  }

  function showLearnMore() {
    openModal("assets/forms/learn-more.jpg", [
      { label: "Back to Intro", action: () => { closeModal(); showIntro(); } },
      { label: "Let’s Play", close: true },
    ]);
  }

  function showHelp(page) {
    const pages = [1, 2, 3, 4, 5, 6];
    const p = pages.includes(page) ? page : 1;
    const buttons = [];
    if (p > 1) {
      buttons.push({
        label: "Back",
        action: () => { closeModal(); showHelp(p - 1); },
      });
    }
    if (p < 6) {
      buttons.push({
        label: "Next",
        action: () => { closeModal(); showHelp(p + 1); },
      });
    }
    buttons.push({ label: "Close", close: true });
    openModal(`assets/forms/help-${p}.jpg`, buttons);
  }

  function showCredits() {
    openModal("assets/forms/credits.jpg", [{ label: "Close", close: true }]);
  }

  // --- Input (document capture so keys work even if focus left a button) ---
  function handleGameKey(e) {
    if (isModalOpen()) {
      if (e.key === "Escape") {
        e.preventDefault();
        if (!nameModal.hidden) finishNameEntry("Player");
        else if (!hsModal.hidden) hideHighscoreModal();
        else closeModal();
      }
      return;
    }
    if (!state.bFormRunning) return;
    if (e.key === "ArrowLeft" || e.key === "Left") {
      e.preventDefault();
      moveImage("Left");
    } else if (e.key === "ArrowRight" || e.key === "Right") {
      e.preventDefault();
      moveImage("Right");
    } else if (e.key === "ArrowUp" || e.key === "Up") {
      e.preventDefault();
      startJump();
    } else if (e.key === "a" || e.key === "A") {
      e.preventDefault();
      moveImage("Left");
    } else if (e.key === "d" || e.key === "D") {
      e.preventDefault();
      moveImage("Right");
    } else if (e.key === "w" || e.key === "W") {
      e.preventDefault();
      startJump();
    } else if (e.key === "Escape") {
      e.preventDefault();
      endGame();
    }
  }

  // Single document-level listener (capture) — do not also bind on #stage
  // or Arrow keys fire twice when the stage is focused.
  document.addEventListener("keydown", handleGameKey, true);

  // Touch / click columns while playing — also refocus for subsequent keys
  stage.addEventListener("pointerdown", (e) => {
    focusStage();
    if (!state.bFormRunning) return;
    const rect = stage.getBoundingClientRect();
    // getBoundingClientRect is already scale-adjusted; map back to logical cols
    const x = ((e.clientX - rect.left) / Math.max(1, rect.width)) * STAGE_W;
    const col = Math.max(0, Math.min(COLS - 1, Math.floor(x / COL_W)));
    if (col < state.baneCol) moveImage("Left");
    else if (col > state.baneCol) moveImage("Right");
  });

  btnStart.addEventListener("click", beginStartSequence);
  btnIntro.addEventListener("click", showIntro);
  btnHelp.addEventListener("click", () => showHelp(1));
  btnCredits.addEventListener("click", showCredits);
  btnHighscore.addEventListener("click", showHighscoreModal);

  // Boot — overlays must stay hidden (CSS display:grid otherwise shows them)
  setShown(overlayGoodLuck, false);
  setShown(overlayGameOver, false);
  setShown(toast, false);
  setupBane();
  setBaneVisible(false); // idle: no Bane until a match starts
  setStatus("Bereit — Start Game · Klicke das Spielfeld, dann ← → ↑");
  fitViewport();
  // Re-fit after images (logo) load — height can change
  window.addEventListener("load", fitViewport);
  const logoImg = document.querySelector(".brand .logo");
  if (logoImg) {
    if (logoImg.complete) fitViewport();
    else logoImg.addEventListener("load", fitViewport);
  }
  console.info("[Strawberry Bane] Port ready. Sounds under assets/sounds/:", Object.values(SOUND_FILES));
})();
