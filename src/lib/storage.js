// storage.js — localStorage with cross-tab sync via BroadcastChannel
const CHANNEL_NAME = 'p25-snags-sync'
let channel = null

try {
  channel = new BroadcastChannel(CHANNEL_NAME)
} catch {}

export function loadSnags(key) {
  try {
    const raw = localStorage.getItem(key)
    return raw ? JSON.parse(raw) : null
  } catch { return null }
}

export function saveSnags(key, data) {
  try {
    const str = JSON.stringify(data)
    localStorage.setItem(key, str)
    channel?.postMessage({ type: 'update', key, data })
  } catch {}
}

export function onSnagsUpdated(callback) {
  if (!channel) return () => {}
  const handler = (e) => {
    if (e.data?.type === 'update') callback(e.data.data)
  }
  channel.addEventListener('message', handler)
  return () => channel.removeEventListener('message', handler)
}
