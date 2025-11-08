const workbookInput = document.querySelector('#workbookInput');
const uploadStatus = document.querySelector('#uploadStatus');
const mappingControls = document.querySelector('#mappingControls');
const questionFieldSelect = document.querySelector('#questionField');
const answerFieldSelect = document.querySelector('#answerField');
const definitionFieldSelect = document.querySelector('#definitionField');
const preview = document.querySelector('#preview');
const previewHead = document.querySelector('#previewHead');
const previewBody = document.querySelector('#previewBody');
const playButton = document.querySelector('#playButton');
const howToButton = document.querySelector('#howToButton');
const howToSection = document.querySelector('#howTo');
const modeCards = document.querySelectorAll('.mode');
const previewQuestionList = document.querySelector('#previewQuestion');
const previewAnswerList = document.querySelector('#previewAnswer');
const gameSection = document.querySelector('#game');
const promptEl = document.querySelector('#prompt');
const choicesEl = document.querySelector('#choices');
const timerEl = document.querySelector('#timer');
const scoreEl = document.querySelector('#score');
const answeredEl = document.querySelector('#answered');
const nextButton = document.querySelector('#nextButton');
const resetButton = document.querySelector('#resetButton');
const summarySection = document.querySelector('#summary');
const summaryMessage = document.querySelector('#summaryMessage');
const playAgainButton = document.querySelector('#playAgainButton');
const choiceTemplate = document.querySelector('#choiceTemplate');

const state = {
  rawRows: [],
  cards: [],
  headers: [],
  mode: 'questionToAnswer',
  timer: null,
  timeRemaining: 60,
  score: 0,
  answered: 0,
  streak: 0,
  round: null,
};

const SCORING = {
  correct: 100,
  incorrect: -25,
};

workbookInput?.addEventListener('change', handleWorkbookUpload);
playButton?.addEventListener('click', startGame);
howToButton?.addEventListener('click', () => toggleSection(howToSection));
playAgainButton?.addEventListener('click', () => {
  summarySection.hidden = true;
  resetGame();
});
resetButton?.addEventListener('click', resetGame);
nextButton?.addEventListener('click', () => {
  if (state.round?.answered) {
    queueNextQuestion();
  }
});

modeCards.forEach((card) => {
  const button = card.querySelector('.mode__select');
  button?.addEventListener('click', () => {
    state.mode = card.dataset.mode;
    modeCards.forEach((c) => c.classList.toggle('mode--active', c === card));
    renderModePreviews();
  });
});

function toggleSection(section) {
  const isHidden = section.hidden;
  howToSection.hidden = true;
  summarySection.hidden = true;
  section.hidden = !isHidden;
}

function handleWorkbookUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  uploadStatus.textContent = '解析中…';
  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target?.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

      if (!json.length) {
        throw new Error('找不到資料列，請確認 Excel 內容。');
      }

      state.headers = Object.keys(json[0]);
      state.rawRows = json;
      state.cards = normalizeCards(state.rawRows);

      populateMappingControls();
      renderPreview(state.rawRows);
      renderModePreviews();

      uploadStatus.textContent = `已載入 ${file.name}，共 ${state.cards.length} 筆題目。`;
      playButton.disabled = false;
    } catch (err) {
      console.error(err);
      uploadStatus.textContent = `⚠️ 讀取失敗：${err.message}`;
      playButton.disabled = true;
      mappingControls.hidden = true;
      preview.hidden = true;
    }
  };

  reader.readAsArrayBuffer(file);
}

function normalizeCards(rows) {
  const fallbackQuestion = state.headers[0];
  const fallbackAnswer = state.headers[1] || fallbackQuestion;

  const questionKey = state.headers.includes(questionFieldSelect.value)
    ? questionFieldSelect.value
    : fallbackQuestion;
  const answerKey = state.headers.includes(answerFieldSelect.value)
    ? answerFieldSelect.value
    : fallbackAnswer;
  const definitionKey = state.headers.includes(definitionFieldSelect.value)
    ? definitionFieldSelect.value
    : '';

  return rows
    .map((row) => {
      const question = String(row?.[questionKey] ?? '').trim();
      const answer = String(row?.[answerKey] ?? '').trim();
      const definition = definitionKey ? String(row?.[definitionKey] ?? '').trim() : '';

      if (!question || !answer) {
        return null;
      }

      return {
        question,
        answer,
        definition: definition || question,
        raw: row,
      };
    })
    .filter(Boolean);
}

function populateMappingControls() {
  const selects = [questionFieldSelect, answerFieldSelect, definitionFieldSelect];
  selects.forEach((select) => {
    if (!select) return;
    select.innerHTML = '<option value="">請選擇</option>';
    state.headers.forEach((header) => {
      const option = document.createElement('option');
      option.value = header;
      option.textContent = header;
      select.append(option);
    });
  });

  mappingControls.hidden = false;

  if (state.headers.length) {
    questionFieldSelect.value = state.headers[0];
    answerFieldSelect.value = state.headers[1] || state.headers[0];
    definitionFieldSelect.value = state.headers[2] || '';
  }

  questionFieldSelect.onchange = updateCards;
  answerFieldSelect.onchange = updateCards;
  definitionFieldSelect.onchange = updateCards;

  updateCards();
}

function updateCards() {
  if (!state.rawRows.length) return;

  state.cards = normalizeCards(state.rawRows);
  renderPreview(state.rawRows.slice(0, 5));
  renderModePreviews();
}

function renderPreview(rows) {
  if (!rows?.length) {
    preview.hidden = true;
    return;
  }

  preview.hidden = false;
  previewHead.innerHTML = '';
  previewBody.innerHTML = '';

  state.headers.forEach((header) => {
    const th = document.createElement('th');
    th.textContent = header;
    previewHead.append(th);
  });

  rows.slice(0, 5).forEach((row) => {
    const tr = document.createElement('tr');
    state.headers.forEach((header) => {
      const td = document.createElement('td');
      td.textContent = String(row?.[header] ?? '');
      tr.append(td);
    });
    previewBody.append(tr);
  });
}

function renderModePreviews() {
  const sample = state.cards.slice(0, 3);

  previewQuestionList.innerHTML = '';
  previewAnswerList.innerHTML = '';

  if (!sample.length) {
    previewQuestionList.innerHTML = '<li>尚無有效題目，請調整欄位或資料。</li>';
    previewAnswerList.innerHTML = '<li>尚無有效題目，請調整欄位或資料。</li>';
    return;
  }

  sample.forEach((card) => {
    const questionItem = document.createElement('li');
    questionItem.textContent = `${card.question} → ${card.answer}`;
    previewQuestionList.append(questionItem);

    const answerItem = document.createElement('li');
    answerItem.textContent = `${card.definition} → ${card.question}`;
    previewAnswerList.append(answerItem);
  });
}

function startGame() {
  if (!state.cards.length) {
    uploadStatus.textContent = '請先匯入題庫。';
    return;
  }

  if (state.cards.length < 2) {
    uploadStatus.textContent = '題庫至少需要兩題，才能產生多個選項。';
    return;
  }

  state.score = 0;
  state.answered = 0;
  state.timeRemaining = 60;
  state.streak = 0;
  summarySection.hidden = true;
  howToSection.hidden = true;
  updateScoreboard();
  queueNextQuestion();
  showGame();
  startTimer();
}

function showGame() {
  gameSection.hidden = false;
  window.scrollTo({ top: gameSection.offsetTop, behavior: 'smooth' });
}

function startTimer() {
  clearInterval(state.timer);
  timerEl.textContent = state.timeRemaining;
  state.timer = setInterval(() => {
    state.timeRemaining -= 1;
    timerEl.textContent = state.timeRemaining;

    if (state.timeRemaining <= 0) {
      endGame();
    }
  }, 1000);
}

function queueNextQuestion() {
  state.round = createRound();
  if (!state.round) {
    endGame();
    return;
  }

  promptEl.textContent = state.round.prompt;
  renderChoices(state.round.options);
  nextButton.disabled = true;
}

function createRound() {
  const cards = [...state.cards];
  if (!cards.length) return null;

  const index = Math.floor(Math.random() * cards.length);
  const card = cards[index];
  const prompt = state.mode === 'questionToAnswer' ? card.question : card.definition;
  const correct = state.mode === 'questionToAnswer' ? card.answer : card.question;

  const uniqueOptions = new Set([correct]);
  while (uniqueOptions.size < 4 && uniqueOptions.size < cards.length) {
    const randomCard = cards[Math.floor(Math.random() * cards.length)];
    const option = state.mode === 'questionToAnswer' ? randomCard.answer : randomCard.question;
    if (option && option !== correct) {
      uniqueOptions.add(option);
    }
  }

  const options = shuffleArray([...uniqueOptions]);

  if (options.length < 2) {
    return null;
  }

  return {
    prompt,
    correct,
    options,
    answered: false,
  };
}

function renderChoices(options) {
  choicesEl.innerHTML = '';
  options.forEach((option) => {
    const node = choiceTemplate.content.cloneNode(true);
    const button = node.querySelector('button');
    button.textContent = option;
    button.addEventListener('click', () => handleChoice(button, option));
    choicesEl.append(node);
  });
}

function handleChoice(button, option) {
  if (!state.round || state.round.answered) return;
  state.round.answered = true;
  const isCorrect = option === state.round.correct;

  if (isCorrect) {
    state.score += SCORING.correct;
    state.streak += 1;
    button.classList.add('choice--correct');
  } else {
    state.score = Math.max(0, state.score + SCORING.incorrect);
    state.streak = 0;
    button.classList.add('choice--incorrect');
  }

  Array.from(choicesEl.querySelectorAll('button')).forEach((choice) => {
    choice.disabled = true;
    if (choice.textContent === state.round.correct) {
      choice.classList.add('choice--correct');
    }
  });

  state.answered += 1;
  updateScoreboard();
  nextButton.disabled = false;
}

function updateScoreboard() {
  scoreEl.textContent = state.score;
  answeredEl.textContent = state.answered;
}

function endGame() {
  clearInterval(state.timer);
  timerEl.textContent = '0';
  summaryMessage.textContent = `你總共回答 ${state.answered} 題，最終得分為 ${state.score} 分！`;
  summarySection.hidden = false;
  gameSection.hidden = true;
}

function resetGame() {
  clearInterval(state.timer);
  state.timeRemaining = 60;
  state.score = 0;
  state.answered = 0;
  state.streak = 0;
  updateScoreboard();
  timerEl.textContent = '60';
  gameSection.hidden = true;
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}
