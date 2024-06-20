let lastTime = false;
function logTime(text) {
  const now = Date.now();
  if (lastTime) {
    const elapsed = now - lastTime;
    console.log(`time: ${elapsed / 1000}s, at ${text}`);
  }
  lastTime = now;
}

function getOrCreate(cacheKey, creationFunction, cacheTime) {
  let cache = CacheService.getScriptCache();
  let value = JSON.parse(cache.get(cacheKey));
  if (value == null) {
    console.log(`No cached ${cacheKey}; creating.`);
    value = creationFunction();
    cache.put(cacheKey, JSON.stringify(value), cacheTime);
  }
  else {
    console.log(`Fetched ${cacheKey} from cache.`);
  }
  return value;
}
