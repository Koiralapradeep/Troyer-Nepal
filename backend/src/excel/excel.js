let locked = false;
const queue = [];

export function withWriteLock(fn) {
  return new Promise((resolve, reject) => {
    queue.push({ fn, resolve, reject });
    runQueue();
  });
}

async function runQueue() {
  if (locked) return;
  const job = queue.shift();
  if (!job) return;

  locked = true;
  try {
    const result = await job.fn();
    job.resolve(result);
  } catch (e) {
    job.reject(e);
  } finally {
    locked = false;
    runQueue();
  }
}