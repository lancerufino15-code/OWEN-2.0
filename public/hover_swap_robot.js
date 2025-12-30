/**
 * Hover/Focus swap widget for the OWEN mascot.
 *
 * Used by: `public/chat.js` to render a static image that swaps to an animated
 * image/video on hover or focus.
 *
 * Key exports:
 * - `createHoverSwapRobot`: DOM helper to build the hover-swap element.
 *
 * Assumptions:
 * - Runs in a browser DOM; media URLs are served as static assets.
 */
const VIDEO_EXTENSIONS = [".mp4", ".webm", ".ogg"];

function isVideoSource(src) {
  if (!src) return false;
  const lower = src.toLowerCase();
  return VIDEO_EXTENSIONS.some(ext => lower.endsWith(ext));
}

function preloadVideo(src) {
  if (!src || document.querySelector(`link[rel="preload"][href="${src}"]`)) return;
  const link = document.createElement("link");
  link.rel = "preload";
  link.as = "video";
  link.href = src;
  document.head.appendChild(link);
}

/**
 * Create a hover-swap robot element with static and animated layers.
 *
 * @param params - Image/video sources and accessibility labels.
 * @returns DOM element ready to insert into the page.
 * @remarks Side effects: preloads media and registers event listeners.
 */
export function createHoverSwapRobot({ staticSrc, animatedSrc, alt, className } = {}) {
  const container = document.createElement("div");
  container.className = `hover-swap-robot${className ? ` ${className}` : ""}`;
  container.tabIndex = 0;
  container.setAttribute("aria-label", alt || "OWEN mascot");

  const frame = document.createElement("div");
  frame.className = "hover-swap-robot__frame";

  const staticImg = document.createElement("img");
  staticImg.src = staticSrc;
  staticImg.alt = alt || "OWEN mascot";
  staticImg.className = "hover-swap-robot__layer hover-swap-robot__layer--static";

  let animatedEl = null;
  const isVideo = isVideoSource(animatedSrc);
  if (isVideo) {
    const video = document.createElement("video");
    video.className = "hover-swap-robot__layer hover-swap-robot__layer--animated is-hidden";
    video.src = animatedSrc;
    video.muted = true;
    video.loop = true;
    video.playsInline = true;
    video.preload = "auto";
    video.setAttribute("aria-hidden", "true");
    animatedEl = video;
    preloadVideo(animatedSrc);
  } else {
    const animatedImg = document.createElement("img");
    animatedImg.src = animatedSrc;
    animatedImg.alt = "";
    animatedImg.className = "hover-swap-robot__layer hover-swap-robot__layer--animated is-hidden";
    animatedImg.setAttribute("aria-hidden", "true");
    animatedEl = animatedImg;
    const preload = new Image();
    preload.src = animatedSrc;
  }

  const showAnimated = () => {
    if (!animatedEl) return;
    animatedEl.classList.remove("is-hidden");
    staticImg.classList.add("is-hidden");
    if (animatedEl.tagName === "VIDEO") {
      animatedEl.play().catch(() => null);
    }
  };

  const showStatic = () => {
    if (!animatedEl) return;
    animatedEl.classList.add("is-hidden");
    staticImg.classList.remove("is-hidden");
    if (animatedEl.tagName === "VIDEO") {
      animatedEl.pause();
      animatedEl.currentTime = 0;
    }
  };

  container.addEventListener("mouseenter", showAnimated);
  container.addEventListener("mouseleave", showStatic);
  container.addEventListener("focus", showAnimated);
  container.addEventListener("blur", showStatic);

  frame.append(staticImg, animatedEl);
  container.append(frame);
  return container;
}
