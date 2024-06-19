let lastTime = false;
function logTime(text) {
  const now = Date.now();
  if (lastTime) {
    const elapsed = now - lastTime;
    console.log(`time: ${elapsed / 1000}s, at ${text}`);
  }
  lastTime = now;
}
