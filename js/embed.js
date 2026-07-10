// Delad init-logik för inbäddade Apps Script-sidor (schema, café-pass m.fl.)
function initEmbed(url, placeholderMarker) {
  const isPlaceholder = !url || url.includes(placeholderMarker);
  const frame = document.getElementById('embed-frame');
  const loading = document.getElementById('embed-loading');
  const fallback = document.getElementById('embed-fallback');
  if (isPlaceholder) {
    frame.style.display = 'none';
    fallback.style.display = 'block';
  } else {
    loading.style.display = 'flex';
    frame.addEventListener('load', () => { loading.style.display = 'none'; });
    frame.src = url;
  }
}
