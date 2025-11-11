const form = document.querySelector('[data-role="upload-form"]');
const fileInput = document.querySelector('[data-role="file-input"]');
const statusBar = document.querySelector('[data-role="status"]');
const errorBox = document.querySelector('[data-role="error"]');
const importedSection = document.querySelector('[data-role="imported-section"]');
const importedBody = document.querySelector('[data-role="imported-body"]');
const parsedSection = document.querySelector('[data-role="parsed-section"]');
const parsedList = document.querySelector('[data-role="parsed-list"]');
const emptyParsed = document.querySelector('[data-role="parsed-empty"]');

function setStatus(text, variant = 'info') {
  if (!statusBar) return;
  statusBar.textContent = text;
  statusBar.dataset.variant = variant;
  statusBar.hidden = false;
}

function clearStatus() {
  if (!statusBar) return;
  statusBar.hidden = true;
  statusBar.textContent = '';
  delete statusBar.dataset.variant;
}

function showError(message) {
  if (!errorBox) return;
  errorBox.textContent = message;
  errorBox.hidden = false;
}

function clearError() {
  if (!errorBox) return;
  errorBox.hidden = true;
  errorBox.textContent = '';
}

function html(strings, ...values) {
  return strings.reduce((result, part, index) => {
    const value = values[index] ?? '';
    return result + part + value;
  }, '');
}

function renderDisplay(display) {
  if (!display) return '';
  switch (display.type) {
    case 'json':
      return html`<pre class="json-block">${display.json}</pre>`;
    case 'list':
      return html`<ul class="value-list">${display.lines
        .map((line) => html`<li>${line}</li>`)
        .join('')}</ul>`;
    case 'link':
    case 'asin':
      return html`<a href="${display.url}" target="_blank" rel="noopener noreferrer">${display.text}</a>`;
    case 'number':
      return html`<span class="number">${display.text}</span>`;
    case 'text':
    default:
      return html`<span>${display.text ?? ''}</span>`;
  }
}

function renderImportedRows(rows) {
  if (!importedSection || !importedBody) return;
  if (!rows || rows.length === 0) {
    importedSection.hidden = true;
    importedBody.innerHTML = '';
    return;
  }
  importedSection.hidden = false;
  importedBody.innerHTML = rows
    .map((row) => {
      let typeLabel = '其它';
      switch (row.display.type) {
        case 'link':
          typeLabel = 'URL';
          break;
        case 'asin':
          typeLabel = 'ASIN';
          break;
        case 'text':
          typeLabel = '文本';
          break;
        case 'number':
          typeLabel = '数值';
          break;
        default:
          typeLabel = '其它';
      }
      return html`<tr>
        <td>${row.row_number}</td>
        <td class="truncate">${renderDisplay(row.display)}</td>
        <td>${typeLabel}</td>
      </tr>`;
    })
    .join('');
}

function renderParsedRows(rows) {
  if (!parsedSection || !parsedList || !emptyParsed) return;
  if (!rows || rows.length === 0) {
    parsedSection.hidden = true;
    emptyParsed.hidden = false;
    parsedList.innerHTML = '';
    return;
  }
  parsedSection.hidden = false;
  emptyParsed.hidden = true;
  parsedList.innerHTML = rows
    .map((row) => {
      return html`
        <article class="product-card">
          <header class="product-card__header">
            <span class="badge">第 ${row.row_number} 行</span>
            ${row.source.type !== 'empty'
              ? html`<span class="source">
                  ${['link', 'asin'].includes(row.source.type)
                    ? html`<a href="${row.source.url}" target="_blank" rel="noopener noreferrer">${row.source.text}</a>`
                    : row.source.text}
                </span>`
              : ''}
          </header>
          <div class="product-card__grid">
            ${row.cells
              .map((cell) => {
                return html`
                  <div class="product-card__cell">
                    <div class="cell-label">${cell.header}</div>
                    <div class="cell-value">${renderDisplay(cell.display)}</div>
                  </div>`;
              })
              .join('')}
          </div>
        </article>`;
    })
    .join('');
}

async function submitForm(event) {
  event.preventDefault();
  if (!form) return;

  const file = fileInput?.files?.[0];
  if (!file) {
    showError('请选择需要上传的 Excel 文件。');
    return;
  }

  clearError();
  setStatus('正在上传文件……');

  const formData = new FormData();
  formData.append('workbook', file);

  try {
    const response = await fetch('/parse', {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      throw new Error(payload.message || '解析失败，请稍后重试。');
    }

    setStatus('文件上传成功，开始解析……');

    // Simulated progress feedback for parsing stages
    const stages = [
      '读取工作簿数据……',
      '识别导入列……',
      '整理解析结果……',
      '渲染界面……',
    ];

    for (const stage of stages) {
      setStatus(stage);
      // eslint-disable-next-line no-await-in-loop
      await new Promise((resolve) => setTimeout(resolve, 280));
    }

    const payload = await response.json();
    if (payload.status !== 'ok') {
      throw new Error(payload.message || '解析失败，请稍后重试。');
    }

    setStatus('解析完成！', 'success');
    renderImportedRows(payload.data.imported_rows);
    renderParsedRows(payload.data.parsed_rows);
  } catch (error) {
    console.error(error);
    showError(error.message || '上传或解析过程中出现问题。');
    setStatus('解析失败', 'error');
    renderImportedRows([]);
    renderParsedRows([]);
  } finally {
    setTimeout(() => {
      clearStatus();
    }, 1600);
  }
}

if (form) {
  form.addEventListener('submit', submitForm);
}
