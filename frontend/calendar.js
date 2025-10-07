const ONE_DAY = 86400000;

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function endOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

function formatISODate(date) {
  const y = date.getFullYear();
  const m = `${date.getMonth() + 1}`.padStart(2, '0');
  const d = `${date.getDate()}`.padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate();
}

function getCalendarGrid(anchorDate) {
  const firstDay = startOfMonth(anchorDate);
  const lastDay = endOfMonth(anchorDate);
  const firstCell = new Date(firstDay);
  firstCell.setDate(firstCell.getDate() - ((firstDay.getDay() + 6) % 7));

  const days = [];
  for (let i = 0; i < 42; i++) {
    const cellDate = new Date(firstCell.getTime() + i * ONE_DAY);
    days.push({
      date: cellDate,
      currentMonth: cellDate.getMonth() === anchorDate.getMonth(),
    });
  }
  return { days, firstDay, lastDay };
}

export function initializeDatePicker(inputId, options = {}) {
  const input = document.getElementById(inputId);
  if (!input) return;

  const {
    minDate,
    maxDate,
    onSelect,
  } = options;

 const parsedDefault = input.value ? new Date(input.value) : new Date();
 let selectedDate = new Date(parsedDefault.getTime());
  let displayDate = new Date(parsedDefault.getFullYear(), parsedDefault.getMonth(), 1);

  const popover = document.createElement('div');
  popover.className = 'fixed inset-0 z-40 hidden';
  popover.innerHTML = `
    <div class="absolute inset-0 bg-black/20"></div>
    <div class="absolute inset-0 flex items-center justify-center px-4">
      <div class="w-full max-w-sm rounded-[28px] bg-white shadow-[0_16px_40px_rgba(0,0,0,0.08)] ring-1 ring-black/5">
        <div class="border-b border-black/5 px-6 py-4 flex items-center justify-between">
          <button class="h-8 w-8 rounded-[14px] border border-neutral-300 bg-neutral-100 text-sm text-neutral-600 transition hover:bg-neutral-200 focus:outline-none focus:ring-2 focus:ring-black/10" data-prev>&lt;</button>
          <div class="text-sm font-medium text-neutral-900" data-title></div>
          <button class="h-8 w-8 rounded-[14px] border border-neutral-300 bg-neutral-100 text-sm text-neutral-600 transition hover:bg-neutral-200 focus:outline-none focus:ring-2 focus:ring-black/10" data-next>&gt;</button>
        </div>
        <div class="px-6 py-4">
          <div class="grid grid-cols-7 gap-2 text-center text-xs font-medium text-neutral-500 mb-3">
            <span>一</span><span>二</span><span>三</span><span>四</span><span>五</span><span class="text-blue-500">六</span><span class="text-blue-500">日</span>
          </div>
          <div class="grid grid-cols-7 gap-2" data-grid></div>
        </div>
        <div class="border-t border-black/5 px-6 py-3 flex justify-end gap-3 text-sm">
          <button class="rounded-[18px] border border-neutral-300 bg-neutral-50 px-4 py-2 text-neutral-600 transition hover:bg-neutral-200 focus:outline-none focus:ring-2 focus:ring-black/10" data-cancel>取消</button>
          <button class="rounded-[18px] border border-neutral-800 bg-[#1d1d1f] px-4 py-2 text-white transition hover:bg-black focus:outline-none focus:ring-2 focus:ring-black/20" data-confirm>完成</button>
        </div>
      </div>
    </div>
  `;
  document.body.appendChild(popover);

  const overlay = popover.querySelector('div');
  const titleEl = popover.querySelector('[data-title]');
  const gridEl = popover.querySelector('[data-grid]');
  const prevBtn = popover.querySelector('[data-prev]');
  const nextBtn = popover.querySelector('[data-next]');
  const confirmBtn = popover.querySelector('[data-confirm]');
  const cancelBtn = popover.querySelector('[data-cancel]');

  function renderCalendar() {
    const { days } = getCalendarGrid(displayDate);
    const formatter = new Intl.DateTimeFormat('zh-Hans', {
      year: 'numeric',
      month: 'long',
    });
    titleEl.textContent = formatter.format(displayDate);

    gridEl.innerHTML = '';
    days.forEach((day) => {
      const isSelected = selectedDate && sameDay(day.date, selectedDate);
      const isToday = sameDay(day.date, new Date());
      const disabled = (minDate && day.date < minDate) || (maxDate && day.date > maxDate);
      const cell = document.createElement('button');
      cell.type = 'button';
      cell.disabled = disabled;
      cell.dataset.date = formatISODate(day.date);
      cell.className = [
        'inline-flex h-8 w-8 items-center justify-center rounded-full text-sm leading-none focus:outline-none select-none transition-colors',
        day.currentMonth ? 'text-neutral-900' : 'text-neutral-400',
        disabled ? 'cursor-not-allowed opacity-30' : 'hover:bg-neutral-200/60',
        isSelected ? 'bg-[#1d1d1f] text-white' : '',
        !isSelected && isToday ? 'border border-[#1d1d1f]/25' : '',
      ].filter(Boolean).join(' ');
      cell.textContent = String(day.date.getDate());
      cell.addEventListener('click', () => {
        selectedDate = new Date(day.date.getTime());
        renderCalendar();
      });
      gridEl.appendChild(cell);
    });
  }

  function showPopover() {
    popover.classList.remove('hidden');
    renderCalendar();
  }

  function hidePopover() {
    popover.classList.add('hidden');
  }

  input.addEventListener('focus', (event) => {
    event.preventDefault();
    showPopover();
  });
  input.addEventListener('click', (event) => {
    event.preventDefault();
    showPopover();
  });

  overlay.addEventListener('click', hidePopover);
  cancelBtn.addEventListener('click', hidePopover);

  confirmBtn.addEventListener('click', () => {
    if (!selectedDate) return;
    const value = formatISODate(selectedDate);
    input.value = value;
    if (typeof onSelect === 'function') {
      onSelect(selectedDate, value);
    }
    hidePopover();
  });

  prevBtn.addEventListener('click', () => {
    displayDate = new Date(displayDate.getFullYear(), displayDate.getMonth() - 1, 1);
    renderCalendar();
  });

  nextBtn.addEventListener('click', () => {
    displayDate = new Date(displayDate.getFullYear(), displayDate.getMonth() + 1, 1);
    renderCalendar();
  });

  // keyboard interactions
  document.addEventListener('keydown', (event) => {
    if (popover.classList.contains('hidden')) return;
    if (event.key === 'Escape') hidePopover();
  });
}

export function setupDatePickers() {
  const startInput = document.getElementById('start-date');
  if (!startInput) return;

  const defaultValue = startInput.value ? new Date(startInput.value) : new Date();
  const minDate = new Date(defaultValue.getTime() - 365 * ONE_DAY);
  const maxDate = new Date(defaultValue.getTime() + 365 * ONE_DAY);

  initializeDatePicker('start-date', {
    minDate,
    maxDate,
    onSelect: (date, value) => {
      startInput.value = value;
    },
  });
}
