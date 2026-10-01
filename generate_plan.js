const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, Header, Footer, PageNumber, LevelFormat
} = require('docx');
const fs = require('fs');

// ─── Цвета ───────────────────────────────────────────────────────────────────
const BLUE       = '1A5DC8';
const LIGHT_BLUE = 'E8F0FE';
const GREEN      = '1E7E34';
const LIGHT_GREEN= 'E6F4EA';
const ORANGE     = 'B45309';
const LIGHT_ORANGE='FFF3E0';
const GRAY_BG    = 'F5F7FA';
const GRAY_TEXT  = '666666';
const BORDER_CLR = 'D0D9E8';
const DARK       = '1A1A2E';

// ─── Хелперы ─────────────────────────────────────────────────────────────────
const cellBorder = (color = BORDER_CLR) => ({
  top:    { style: BorderStyle.SINGLE, size: 1, color },
  bottom: { style: BorderStyle.SINGLE, size: 1, color },
  left:   { style: BorderStyle.SINGLE, size: 1, color },
  right:  { style: BorderStyle.SINGLE, size: 1, color },
});

const noBorder = () => ({
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
});

const cell = (children, opts = {}) => new TableCell({
  borders: opts.borders ?? cellBorder(),
  width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
  shading: opts.bg ? { fill: opts.bg, type: ShadingType.CLEAR } : undefined,
  margins: { top: 100, bottom: 100, left: 140, right: 140 },
  verticalAlign: opts.va ?? VerticalAlign.TOP,
  children,
});

const p = (text, opts = {}) => new Paragraph({
  alignment: opts.align ?? AlignmentType.LEFT,
  spacing: { before: opts.before ?? 0, after: opts.after ?? 0 },
  children: [new TextRun({
    text,
    font: 'Arial',
    size: opts.size ?? 22,
    bold: opts.bold ?? false,
    color: opts.color ?? DARK,
    italics: opts.italic ?? false,
  })],
});

const space = (pt = 120) => new Paragraph({
  spacing: { before: 0, after: 0, line: pt },
  children: [new TextRun({ text: '', font: 'Arial', size: 22 })],
});

// ─── Нумерованные списки ─────────────────────────────────────────────────────
const numbering = {
  config: [
    {
      reference: 'steps',
      levels: [{
        level: 0,
        format: LevelFormat.DECIMAL,
        text: '%1.',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 580, hanging: 340 } } }
      }]
    },
    {
      reference: 'bullets',
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: '\u2022',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 500, hanging: 280 } } }
      }]
    }
  ]
};

const listItem = (text, ref = 'bullets', color = DARK, size = 22) =>
  new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: 'Arial', size, color })]
  });

// ─── Таблица: шаги реализации ─────────────────────────────────────────────────
function stepRow(num, title, desc, time, bg = 'FFFFFF') {
  return new TableRow({
    children: [
      cell([p(num, { bold: true, color: BLUE, size: 24, before: 60, after: 60 })],
        { width: 700, bg, borders: cellBorder() }),
      cell([
        p(title, { bold: true, size: 22, color: DARK, before: 60 }),
        p(desc, { size: 20, color: GRAY_TEXT, before: 40, after: 60 }),
      ], { width: 5800, bg, borders: cellBorder() }),
      cell([p(time, { size: 20, color: GRAY_TEXT, before: 60, after: 60, align: AlignmentType.CENTER })],
        { width: 1600, bg, borders: cellBorder(), va: VerticalAlign.CENTER }),
    ]
  });
}

// ─── Таблица: логика заполнения ───────────────────────────────────────────────
function logicRow(input, arrow, output, inputBg, outputBg) {
  return new TableRow({
    children: [
      cell([p(input, { bold: true, size: 21, before: 80, after: 80 })],
        { width: 3000, bg: inputBg, borders: cellBorder() }),
      cell([p(arrow, { size: 22, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })],
        { width: 600, bg: 'FFFFFF', borders: cellBorder() }),
      cell([p(output, { size: 21, color: GRAY_TEXT, before: 80, after: 80 })],
        { width: 4560, bg: outputBg, borders: cellBorder() }),
    ]
  });
}

// ─── Документ ────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering,
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 22 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 6 } },
            spacing: { after: 0 },
            children: [
              new TextRun({ text: 'Автоматизация географии в Битрикс24', font: 'Arial', size: 18, color: GRAY_TEXT }),
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: BORDER_CLR, space: 6 } },
            spacing: { before: 80 },
            children: [
              new TextRun({ text: 'Стр. ', font: 'Arial', size: 18, color: GRAY_TEXT }),
              new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 18, color: GRAY_TEXT }),
            ]
          })
        ]
      })
    },
    children: [

      // ── Заголовок ────────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 100 },
        children: [
          new TextRun({ text: 'Автоматизация географии сделок', font: 'Arial', size: 40, bold: true, color: DARK })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
        children: [
          new TextRun({ text: 'в CRM Битрикс24', font: 'Arial', size: 30, bold: false, color: BLUE })
        ]
      }),

      space(180),

      // ── Раздел 1: Цель ────────────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: '1. Цель задачи', font: 'Arial', size: 26, bold: true, color: DARK })]
      }),

      new Table({
        width: { size: 9506, type: WidthType.DXA },
        columnWidths: [9506],
        rows: [new TableRow({ children: [
          cell([
            p('Добавить в карточку сделки три поля для географии проектов, чтобы данные корректно передавались в BI-систему (Power BI) через выгрузку из Битрикс24.', { size: 22, before: 60, after: 80 }),
            new Paragraph({
              spacing: { before: 0, after: 60 },
              children: [
                new TextRun({ text: 'Три поля:', font: 'Arial', size: 22, bold: true, color: DARK }),
              ]
            }),
            listItem('Федеральный округ (ЦФО, СЗФО, ЮФО, СКФО, ПФО, УФО, СФО, ДФО)', 'bullets'),
            listItem('Область / Регион', 'bullets'),
            listItem('Населённый пункт', 'bullets'),
          ], { bg: LIGHT_BLUE, borders: cellBorder(BLUE) })
        ]})]
      }),

      space(240),

      // ── Раздел 2: Ключевой принцип ────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: '2. Ключевой принцип работы', font: 'Arial', size: 26, bold: true, color: DARK })]
      }),

      p('Менеджер заполняет то, что знает — система автоматически дополняет остальное.', { size: 22, color: GRAY_TEXT, before: 0, after: 120 }),

      new Table({
        width: { size: 9506, type: WidthType.DXA },
        columnWidths: [3000, 600, 4560, 1346],
        rows: [
          // Заголовок
          new TableRow({
            tableHeader: true,
            children: [
              cell([p('Что вводит менеджер', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80 })],
                { width: 3000, bg: BLUE, borders: cellBorder(BLUE) }),
              cell([p('', { size: 20 })],
                { width: 600, bg: BLUE, borders: cellBorder(BLUE) }),
              cell([p('Что подтягивается автоматически', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80 })],
                { width: 4560, bg: BLUE, borders: cellBorder(BLUE) }),
              cell([p('Города?', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80, align: AlignmentType.CENTER })],
                { width: 1346, bg: BLUE, borders: cellBorder(BLUE) }),
            ]
          }),
          logicRow('Федеральный округ', '→', 'Все регионы этого ФО', LIGHT_BLUE, LIGHT_GREEN),
          new TableRow({ children: [
            cell([p('', { size: 20 })], { width: 3000, bg: LIGHT_BLUE, borders: cellBorder() }),
            cell([p('', { size: 20 })], { width: 600, bg: 'FFFFFF', borders: cellBorder() }),
            cell([p('', { size: 20 })], { width: 4560, bg: LIGHT_GREEN, borders: cellBorder() }),
            cell([p('Нет', { size: 21, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })],
              { width: 1346, bg: GRAY_BG, borders: cellBorder() }),
          ]}),
          logicRow('Регион', '→', 'Федеральный округ', LIGHT_GREEN, LIGHT_BLUE),
          new TableRow({ children: [
            cell([p('', { size: 20 })], { width: 3000, bg: LIGHT_GREEN, borders: cellBorder() }),
            cell([p('', { size: 20 })], { width: 600, bg: 'FFFFFF', borders: cellBorder() }),
            cell([p('', { size: 20 })], { width: 4560, bg: LIGHT_BLUE, borders: cellBorder() }),
            cell([p('Нет', { size: 21, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })],
              { width: 1346, bg: GRAY_BG, borders: cellBorder() }),
          ]}),
          logicRow('Населённый пункт', '→', 'Регион + Федеральный округ', LIGHT_ORANGE, LIGHT_GREEN),
          new TableRow({ children: [
            cell([p('', { size: 20 })], { width: 3000, bg: LIGHT_ORANGE, borders: cellBorder() }),
            cell([p('', { size: 20 })], { width: 600, bg: 'FFFFFF', borders: cellBorder() }),
            cell([p('', { size: 20 })], { width: 4560, bg: LIGHT_GREEN, borders: cellBorder() }),
            cell([p('Да', { size: 21, color: GREEN, bold: true, before: 80, after: 80, align: AlignmentType.CENTER })],
              { width: 1346, bg: GRAY_BG, borders: cellBorder() }),
          ]}),
        ]
      }),

      space(160),

      // Важно: удаление
      new Table({
        width: { size: 9506, type: WidthType.DXA },
        columnWidths: [9506],
        rows: [new TableRow({ children: [
          cell([
            p('Удаление — только на своём уровне, без каскада:', { size: 22, bold: true, before: 60, after: 80 }),
            listItem('Удалил регион → ФО остаётся', 'bullets', GRAY_TEXT),
            listItem('Удалил город → регион и ФО остаются', 'bullets', GRAY_TEXT),
            listItem('Регионы внутри ФО можно удалять по одному — сам ФО при этом не исчезает', 'bullets', GRAY_TEXT),
          ], { bg: LIGHT_GREEN, borders: cellBorder(GREEN) })
        ]})]
      }),

      space(240),

      // ── Раздел 3: Интерфейс ───────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: '3. Интерфейс виджета в карточке сделки', font: 'Arial', size: 26, bold: true, color: DARK })]
      }),

      new Table({
        width: { size: 9506, type: WidthType.DXA },
        columnWidths: [3100, 6406],
        rows: [
          new TableRow({ tableHeader: true, children: [
            cell([p('Блок', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80 })], { width: 3100, bg: BLUE, borders: cellBorder(BLUE) }),
            cell([p('Описание', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80 })], { width: 6406, bg: BLUE, borders: cellBorder(BLUE) }),
          ]}),
          new TableRow({ children: [
            cell([p('Федеральный округ', { bold: true, size: 21, before: 80, after: 80 })], { width: 3100, bg: LIGHT_BLUE, borders: cellBorder() }),
            cell([p('8 кнопок ФО — нажал на нужный, все его регионы появились автоматически. Регионы внутри группы можно удалять по одному.', { size: 21, before: 80, after: 80 })], { width: 6406, borders: cellBorder() }),
          ]}),
          new TableRow({ children: [
            cell([p('Область / Регион', { bold: true, size: 21, before: 80, after: 80 })], { width: 3100, bg: LIGHT_GREEN, borders: cellBorder() }),
            cell([p('Поиск с автокомплитом. Регионы отображаются сгруппированными под своим ФО — сразу видно из какого округа что удалять.', { size: 21, before: 80, after: 80 })], { width: 6406, borders: cellBorder() }),
          ]}),
          new TableRow({ children: [
            cell([p('Населённый пункт', { bold: true, size: 21, before: 80, after: 80 })], { width: 3100, bg: LIGHT_ORANGE, borders: cellBorder() }),
            cell([p('Поиск с автокомплитом по 826 городам РФ. Мультивыбор. При вводе список сужается по введённым буквам.', { size: 21, before: 80, after: 80 })], { width: 6406, borders: cellBorder() }),
          ]}),
        ]
      }),

      space(240),

      // ── Раздел 4: Архитектура ─────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: '4. Техническая архитектура', font: 'Arial', size: 26, bold: true, color: DARK })]
      }),

      p('Решение не требует платных серверов. Все компоненты — бесплатные или встроенные в Битрикс24.', { size: 22, color: GRAY_TEXT, before: 0, after: 140 }),

      new Table({
        width: { size: 9506, type: WidthType.DXA },
        columnWidths: [2600, 3706, 3200],
        rows: [
          new TableRow({ tableHeader: true, children: [
            cell([p('Компонент', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80 })], { width: 2600, bg: BLUE, borders: cellBorder(BLUE) }),
            cell([p('Что это', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80 })], { width: 3706, bg: BLUE, borders: cellBorder(BLUE) }),
            cell([p('Стоимость', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80 })], { width: 3200, bg: BLUE, borders: cellBorder(BLUE) }),
          ]}),
          new TableRow({ children: [
            cell([p('База городов РФ', { bold: true, size: 21, before: 80, after: 80 })], { width: 2600, bg: GRAY_BG, borders: cellBorder() }),
            cell([p('826 городов с привязкой регион → ФО. Файл JSON на GitHub Pages.', { size: 21, before: 80, after: 80 })], { width: 3706, borders: cellBorder() }),
            cell([p('Бесплатно', { size: 21, color: GREEN, bold: true, before: 80, after: 80 })], { width: 3200, borders: cellBorder() }),
          ]}),
          new TableRow({ children: [
            cell([p('Виджет (HTML/JS)', { bold: true, size: 21, before: 80, after: 80 })], { width: 2600, bg: GRAY_BG, borders: cellBorder() }),
            cell([p('Страница с интерфейсом, встраивается в карточку сделки Б24 через iframe.', { size: 21, before: 80, after: 80 })], { width: 3706, borders: cellBorder() }),
            cell([p('Бесплатно (GitHub Pages)', { size: 21, color: GREEN, bold: true, before: 80, after: 80 })], { width: 3200, borders: cellBorder() }),
          ]}),
          new TableRow({ children: [
            cell([p('Битрикс24 Local App', { bold: true, size: 21, before: 80, after: 80 })], { width: 2600, bg: GRAY_BG, borders: cellBorder() }),
            cell([p('Приложение в Б24, которое встраивает виджет в карточку сделки.', { size: 21, before: 80, after: 80 })], { width: 3706, borders: cellBorder() }),
            cell([p('Встроено в Б24', { size: 21, color: GREEN, bold: true, before: 80, after: 80 })], { width: 3200, borders: cellBorder() }),
          ]}),
          new TableRow({ children: [
            cell([p('REST API + Webhook', { bold: true, size: 21, before: 80, after: 80 })], { width: 2600, bg: GRAY_BG, borders: cellBorder() }),
            cell([p('Запись данных из виджета в поля сделки через входящий вебхук Б24.', { size: 21, before: 80, after: 80 })], { width: 3706, borders: cellBorder() }),
            cell([p('Встроено в Б24', { size: 21, color: GREEN, bold: true, before: 80, after: 80 })], { width: 3200, borders: cellBorder() }),
          ]}),
        ]
      }),

      space(240),

      // ── Раздел 5: Шаги реализации ─────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: '5. Шаги реализации', font: 'Arial', size: 26, bold: true, color: DARK })]
      }),

      new Table({
        width: { size: 9506, type: WidthType.DXA },
        columnWidths: [700, 5800, 1600, 1406],
        rows: [
          new TableRow({ tableHeader: true, children: [
            cell([p('#', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80 })], { width: 700, bg: BLUE, borders: cellBorder(BLUE) }),
            cell([p('Шаг', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80 })], { width: 5800, bg: BLUE, borders: cellBorder(BLUE) }),
            cell([p('Время', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1600, bg: BLUE, borders: cellBorder(BLUE) }),
            cell([p('Кто', { bold: true, size: 20, color: 'FFFFFF', before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1406, bg: BLUE, borders: cellBorder(BLUE) }),
          ]}),
          new TableRow({ children: [
            cell([p('1', { bold: true, color: BLUE, size: 24, before: 80, after: 80 })], { width: 700, bg: LIGHT_BLUE, borders: cellBorder() }),
            cell([
              p('Предоставление прав администратора Б24', { bold: true, size: 21, before: 80 }),
              p('Единоразово на время настройки. После завершения права возвращаются к стандартным.', { size: 20, color: GRAY_TEXT, before: 40, after: 80 }),
            ], { width: 5800, borders: cellBorder() }),
            cell([p('1–2 дня', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1600, borders: cellBorder(), va: VerticalAlign.CENTER }),
            cell([p('IT отдел', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1406, borders: cellBorder(), va: VerticalAlign.CENTER }),
          ]}),
          new TableRow({ children: [
            cell([p('2', { bold: true, color: BLUE, size: 24, before: 80, after: 80 })], { width: 700, bg: LIGHT_BLUE, borders: cellBorder() }),
            cell([
              p('Создание входящего вебхука в Б24', { bold: true, size: 21, before: 80 }),
              p('Настройки → Разработчикам → Входящий вебхук. Нужен для записи данных в сделки через REST API.', { size: 20, color: GRAY_TEXT, before: 40, after: 80 }),
            ], { width: 5800, borders: cellBorder() }),
            cell([p('10 мин', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1600, borders: cellBorder(), va: VerticalAlign.CENTER }),
            cell([p('Разработчик', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1406, borders: cellBorder(), va: VerticalAlign.CENTER }),
          ]}),
          new TableRow({ children: [
            cell([p('3', { bold: true, color: BLUE, size: 24, before: 80, after: 80 })], { width: 700, bg: LIGHT_BLUE, borders: cellBorder() }),
            cell([
              p('Создание 3 пользовательских полей в сделках CRM', { bold: true, size: 21, before: 80 }),
              p('Через REST API добавляются поля: Федеральный округ, Область/Регион, Населённый пункт. Тип — строка.', { size: 20, color: GRAY_TEXT, before: 40, after: 80 }),
            ], { width: 5800, borders: cellBorder() }),
            cell([p('15 мин', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1600, borders: cellBorder(), va: VerticalAlign.CENTER }),
            cell([p('Разработчик', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1406, borders: cellBorder(), va: VerticalAlign.CENTER }),
          ]}),
          new TableRow({ children: [
            cell([p('4', { bold: true, color: BLUE, size: 24, before: 80, after: 80 })], { width: 700, bg: LIGHT_BLUE, borders: cellBorder() }),
            cell([
              p('Размещение виджета на GitHub Pages', { bold: true, size: 21, before: 80 }),
              p('Загрузка файлов виджета (HTML, JS, база городов) на бесплатный хостинг. Сторонний сервер не нужен.', { size: 20, color: GRAY_TEXT, before: 40, after: 80 }),
            ], { width: 5800, borders: cellBorder() }),
            cell([p('20 мин', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1600, borders: cellBorder(), va: VerticalAlign.CENTER }),
            cell([p('Разработчик', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1406, borders: cellBorder(), va: VerticalAlign.CENTER }),
          ]}),
          new TableRow({ children: [
            cell([p('5', { bold: true, color: BLUE, size: 24, before: 80, after: 80 })], { width: 700, bg: LIGHT_BLUE, borders: cellBorder() }),
            cell([
              p('Регистрация Local App в Б24 и встройка виджета', { bold: true, size: 21, before: 80 }),
              p('Разработчикам → Приложения → добавить приложение. Виджет встраивается в карточку сделки через iframe и доступен во всех сделках.', { size: 20, color: GRAY_TEXT, before: 40, after: 80 }),
            ], { width: 5800, borders: cellBorder() }),
            cell([p('30 мин', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1600, borders: cellBorder(), va: VerticalAlign.CENTER }),
            cell([p('Разработчик', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1406, borders: cellBorder(), va: VerticalAlign.CENTER }),
          ]}),
          new TableRow({ children: [
            cell([p('6', { bold: true, color: BLUE, size: 24, before: 80, after: 80 })], { width: 700, bg: LIGHT_BLUE, borders: cellBorder() }),
            cell([
              p('Финальная настройка и тестирование', { bold: true, size: 21, before: 80 }),
              p('Подключение вебхука и кодов полей в код виджета. Проверка менеджером в реальной сделке.', { size: 20, color: GRAY_TEXT, before: 40, after: 80 }),
            ], { width: 5800, borders: cellBorder() }),
            cell([p('30 мин', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1600, borders: cellBorder(), va: VerticalAlign.CENTER }),
            cell([p('Разработчик', { size: 20, color: GRAY_TEXT, before: 80, after: 80, align: AlignmentType.CENTER })], { width: 1406, borders: cellBorder(), va: VerticalAlign.CENTER }),
          ]}),
        ]
      }),

      space(240),

      // ── Раздел 6: Итог ────────────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({ text: '6. Итог', font: 'Arial', size: 26, bold: true, color: DARK })]
      }),

      new Table({
        width: { size: 9506, type: WidthType.DXA },
        columnWidths: [4753, 4753],
        rows: [
          new TableRow({ children: [
            cell([
              p('Технической работы', { bold: true, size: 22, color: BLUE, before: 100, after: 60 }),
              p('~1.5 часа', { bold: true, size: 32, color: DARK, before: 0, after: 100 }),
            ], { width: 4753, bg: LIGHT_BLUE, borders: cellBorder(BLUE) }),
            cell([
              p('Ожидание доступов от IT', { bold: true, size: 22, color: GRAY_TEXT, before: 100, after: 60 }),
              p('1–2 рабочих дня', { bold: true, size: 32, color: DARK, before: 0, after: 100 }),
            ], { width: 4753, bg: GRAY_BG, borders: cellBorder() }),
          ]}),
          new TableRow({ children: [
            new TableCell({
              columnSpan: 2,
              borders: cellBorder(GREEN),
              width: { size: 9506, type: WidthType.DXA },
              shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 100, left: 140, right: 140 },
              children: [
                p('Стоимость внешней инфраструктуры', { bold: true, size: 22, color: GREEN, before: 100, after: 60 }),
                p('0 ₽ — только встроенные инструменты Б24 и бесплатный хостинг', { size: 21, color: DARK, before: 0, after: 100 }),
              ]
            }),
          ]})
        ]
      }),

      space(200),
      p('После получения прав администратора — всё остальное реализуется за один рабочий сеанс.', { size: 21, color: GRAY_TEXT, italic: true }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('C:\\Claude\\География сделок\\План автоматизации географии Битрикс24.docx', buf);
  console.log('OK');
});
