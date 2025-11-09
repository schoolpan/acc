const cardElement = document.getElementById('card');
const frontFace = cardElement.querySelector('.card__face--front');
const backFace = cardElement.querySelector('.card__face--back');
const fileInput = document.getElementById('fileInput');
const statusElement = document.getElementById('status');
const progressElement = document.getElementById('progress');
const playButton = document.getElementById('playButton');
const pauseButton = document.getElementById('pauseButton');
const randomButton = document.getElementById('randomButton');
const prevButton = document.getElementById('prevButton');
const nextButton = document.getElementById('nextButton');
const directionSelect = document.getElementById('directionSelect');
const intervalInput = document.getElementById('intervalInput');

let cards = [];
let currentIndex = 0;
let isPlaying = false;
let playTimer = null;
let flipTimer = null;

const ORIENTATION = {
  TERM_TO_DEFINITION: 'term-to-definition',
  DEFINITION_TO_TERM: 'definition-to-term',
  MIXED: 'mixed'
};

function resetState() {
  clearTimers();
  isPlaying = false;
  cards = [];
  currentIndex = 0;
  updateProgress();
  updateCardFaces('', '');
  statusElement.textContent = '請先匯入 Excel 檔案 (第一欄詞語，第二欄定義)。';
}

function clearTimers() {
  if (playTimer) {
    clearInterval(playTimer);
    playTimer = null;
  }
  if (flipTimer) {
    clearTimeout(flipTimer);
    flipTimer = null;
  }
}

function updateProgress() {
  progressElement.textContent = cards.length ? `${currentIndex + 1} / ${cards.length}` : '0 / 0';
}

function updateCardFaces(frontText, backText) {
  frontFace.textContent = frontText || '準備開始學習！';
  backFace.textContent = backText || '匯入資料後即可開始。';
  cardElement.classList.remove('is-flipped');
}

function setStatus(message, isError = false) {
  statusElement.textContent = message;
  statusElement.style.color = isError ? '#c92a2a' : '#666';
}

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
        resolve(json);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function parseRows(rows) {
  const parsed = [];
  const headerKeywords = ['詞', 'term', '單字', 'definition', '定義', 'meaning', '翻譯'];

  rows.forEach((row, index) => {
    if (
      index === 0 &&
      typeof row[0] === 'string' &&
      typeof row[1] === 'string' &&
      headerKeywords.some(keyword => row[0].toLowerCase().includes(keyword) || row[1].toLowerCase().includes(keyword))
    ) {
      // Skip header row that looks like titles.
      return;
    }

    const term = (row?.[0] ?? '').toString().trim();
    const definition = (row?.[1] ?? '').toString().trim();

    if (term && definition) {
      parsed.push({ term, definition });
    }
  });

  return parsed;
}

function shuffle(array) {
  const result = [...array];
  for (let i = result.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [result[i], result[j]] = [result[j], result[i]];
  }
  return result;
}

function currentOrientation() {
  const value = directionSelect.value;
  if (value === ORIENTATION.MIXED) {
    return Math.random() > 0.5 ? ORIENTATION.TERM_TO_DEFINITION : ORIENTATION.DEFINITION_TO_TERM;
  }
  return value;
}

function renderCard(index = currentIndex, keepFlipState = false) {
  if (!cards.length) {
    updateCardFaces('', '');
    return;
  }

  const card = cards[index];
  const orientation = currentOrientation();
  const front = orientation === ORIENTATION.TERM_TO_DEFINITION ? card.term : card.definition;
  const back = orientation === ORIENTATION.TERM_TO_DEFINITION ? card.definition : card.term;

  if (!keepFlipState) {
    cardElement.classList.remove('is-flipped');
  }

  frontFace.textContent = front;
  backFace.textContent = back;
  updateProgress();
}

function goTo(index) {
  if (!cards.length) return;
  currentIndex = (index + cards.length) % cards.length;
  renderCard();
}

function nextCard() {
  goTo(currentIndex + 1);
}

function prevCard() {
  goTo(currentIndex - 1);
}

function toggleCard() {
  cardElement.classList.toggle('is-flipped');
}

function startAutoPlay() {
  if (!cards.length) {
    setStatus('請先匯入卡片資料。', true);
    return;
  }

  clearTimers();
  isPlaying = true;
  setStatus('自動播放中…');
  const interval = Math.max(Number(intervalInput.value) || 6, 2) * 1000;
  const flipDelay = interval / 2;

  const showCard = index => {
    if (flipTimer) {
      clearTimeout(flipTimer);
      flipTimer = null;
    }
    currentIndex = index;
    renderCard();
    flipTimer = setTimeout(() => {
      toggleCard();
    }, flipDelay);
  };

  showCard(currentIndex);

  playTimer = setInterval(() => {
    const nextIndex = (currentIndex + 1) % cards.length;
    showCard(nextIndex);
  }, interval);
}

function pauseAutoPlay() {
  if (!isPlaying) return;
  clearTimers();
  isPlaying = false;
  setStatus('已暫停自動播放。');
}

function randomizeCards() {
  if (!cards.length) return;
  pauseAutoPlay();
  cards = shuffle(cards);
  currentIndex = 0;
  renderCard();
  setStatus('已打散卡片順序。');
}

function loadCardsFromRows(rows) {
  const parsed = parseRows(rows);
  if (!parsed.length) {
    setStatus('找不到可用的詞語資料，請確認檔案內容。', true);
    return;
  }

  cards = parsed;
  currentIndex = 0;
  renderCard();
  updateProgress();
  setStatus(`成功匯入 ${cards.length} 張卡片！`);
}

fileInput.addEventListener('change', async event => {
  const [file] = event.target.files;
  if (!file) return;

  try {
    pauseAutoPlay();
    setStatus(`正在讀取 ${file.name} …`);
    const rows = await readExcel(file);
    loadCardsFromRows(rows);
  } catch (error) {
    console.error(error);
    setStatus('讀取檔案時發生錯誤，請確認格式是否正確。', true);
  }
});

cardElement.addEventListener('click', () => {
  if (!cards.length) return;
  toggleCard();
});

cardElement.addEventListener('keyup', event => {
  if (event.code === 'Space' || event.code === 'Enter') {
    toggleCard();
  }
});

playButton.addEventListener('click', () => {
  startAutoPlay();
});

pauseButton.addEventListener('click', () => {
  pauseAutoPlay();
});

randomButton.addEventListener('click', () => {
  randomizeCards();
});

nextButton.addEventListener('click', () => {
  pauseAutoPlay();
  nextCard();
});

prevButton.addEventListener('click', () => {
  pauseAutoPlay();
  prevCard();
});

directionSelect.addEventListener('change', () => {
  renderCard(currentIndex, true);
});

document.addEventListener('keydown', event => {
  if (event.target instanceof HTMLInputElement || event.target instanceof HTMLSelectElement) {
    return;
  }

  switch (event.key) {
    case 'ArrowRight':
      pauseAutoPlay();
      nextCard();
      break;
    case 'ArrowLeft':
      pauseAutoPlay();
      prevCard();
      break;
    case ' ': // space
      event.preventDefault();
      toggleCard();
      break;
    default:
      break;
  }
});

window.addEventListener('blur', () => {
  pauseAutoPlay();
});

resetState();
