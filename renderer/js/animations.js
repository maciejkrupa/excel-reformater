const runBtn = document.getElementById('runBtn');
const prepBtn = document.getElementById('prepBtn');
const openBtn = document.getElementById('openBtn');
const msgOne = document.getElementById('msgOne');
const msgTwo = document.getElementById('msgTwo');
const msgThree = document.getElementById('msgThree');
const roller = document.getElementById('roller');
const fadescreen = document.getElementById('fadescreen');

let prepToRun = function() {
  prepBtn.classList.remove('bigBtn--visible');
  prepBtn.classList.add('hidden');
  runBtn.classList.remove('hidden');
  runBtn.classList.add('bigBtn--visible');
  msgTwo.classList.remove('hidden');
  msgTwo.classList.add('container--visible');
  msgOne.style.opacity = '0.5';
};

let runApp = function() {
  roller.classList.remove('hidden');
  roller.classList.add('roller--visible');
  fadescreen.classList.remove('hidden');
  fadescreen.classList.add('fadescreen--visible');
  msgThree.classList.remove('container--visible');
  msgThree.classList.add('hidden');
  fadescreen.style.transition = 'opacity 0.2s ease-out';
};

let stopApp = function() {
  roller.classList.remove('roller--visible');
  roller.classList.add('hidden');
  fadescreen.classList.remove('fadescreen--visible');
  fadescreen.classList.add('hidden');
  msgThree.classList.remove('hidden');
  msgThree.classList.add('container--visible');
  msgTwo.style.opacity = '0.5';
  runBtn.classList.remove('bigBtn--visible');
  runBtn.classList.add('hidden');
  openBtn.classList.remove('hidden');
  openBtn.classList.add('bigBtn--visible');
};

function exportAnimations() {
  exports.prepToRun = prepToRun;
  exports.runApp = runApp;
  exports.stopApp = stopApp;
};

exportAnimations();
