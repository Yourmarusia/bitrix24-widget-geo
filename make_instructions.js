const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  ExternalHyperlink, LevelFormat, Header, Footer, PageNumber,
  UnderlineType
} = require('docx');
const fs = require('fs');

// ── Цвета ────────────────────────────────────────────────────────────────────
const BLUE       = '2F75B6';
const LIGHT_BLUE = 'DCE6F1';
const GREEN      = '375623';
const LIGHT_GREEN = 'E2EFDA';
const GREY_BG    = 'F2F2F2';
const DARK_TEXT  = '1F2937';

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// ── Вспомогательные функции ──────────────────────────────────────────────────
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, bold: true, size: 32, color: BLUE, font: 'Arial' })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, color: '1F497D', font: 'Arial' })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, size: 22, font: 'Arial', color: DARK_TEXT, ...opts })]
  });
}

function li(text, num = false, ref = num ? 'nums' : 'bullets') {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22, font: 'Arial', color: DARK_TEXT })]
  });
}

function tip(title, text) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    margins: { top: 80, bottom: 80 },
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: 9026, type: WidthType.DXA },
        shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        children: [
          new Paragraph({ spacing: { before: 0, after: 60 }, children: [
            new TextRun({ text: '💡 ' + title, bold: true, size: 22, font: 'Arial', color: GREEN })
          ]}),
          new Paragraph({ spacing: { before: 0, after: 0 }, children: [
            new TextRun({ text, size: 21, font: 'Arial', color: DARK_TEXT })
          ]})
        ]
      })]
    })]
  });
}

function stepRow(num, title, desc) {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 900, type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 120, right: 120 },
        verticalAlign: 'center',
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: String(num), bold: true, size: 28, color: 'FFFFFF', font: 'Arial' })]
        })]
      }),
      new TableCell({
        borders,
        width: { size: 8126, type: WidthType.DXA },
        shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        children: [
          new Paragraph({ spacing: { before: 0, after: 40 }, children: [
            new TextRun({ text: title, bold: true, size: 23, font: 'Arial', color: '1F497D' })
          ]}),
          new Paragraph({ spacing: { before: 0, after: 0 }, children: [
            new TextRun({ text: desc, size: 21, font: 'Arial', color: DARK_TEXT })
          ]})
        ]
      })
    ]
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LIGHT_BLUE, space: 1 } },
    children: []
  });
}

function empty(before = 120) {
  return new Paragraph({ spacing: { before, after: 0 }, children: [] });
}

// ── Таблица логики автозаполнения ─────────────────────────────────────────────
function logicTable() {
  const headerRow = new TableRow({
    tableHeader: true,
    children: [
      ['Что вводите', 2800],
      ['Что подтянется автоматически', 6226],
    ].map(([text, w]) => new TableCell({
      borders,
      width: { size: w, type: WidthType.DXA },
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      children: [new Paragraph({ children: [
        new TextRun({ text, bold: true, size: 21, font: 'Arial', color: 'FFFFFF' })
      ]})]
    }))
  });

  const dataRows = [
    ['ФО (кнопка)', 'Все регионы этого федерального округа'],
    ['Регион (поиск)', 'Федеральный округ'],
    ['Город (поиск)', 'Регион + Федеральный округ'],
  ].map(([col1, col2], i) => new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 2800, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? GREY_BG : 'FFFFFF', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 140, right: 140 },
        children: [new Paragraph({ children: [
          new TextRun({ text: col1, bold: true, size: 21, font: 'Arial', color: DARK_TEXT })
        ]})]
      }),
      new TableCell({
        borders,
        width: { size: 6226, type: WidthType.DXA },
        shading: { fill: i % 2 === 0 ? GREY_BG : 'FFFFFF', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 140, right: 140 },
        children: [new Paragraph({ children: [
          new TextRun({ text: col2, size: 21, font: 'Arial', color: DARK_TEXT })
        ]})]
      }),
    ]
  }));

  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2800, 6226],
    rows: [headerRow, ...dataRows]
  });
}

// ── Основной документ ─────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 22, color: DARK_TEXT } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: 'Arial', color: BLUE },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: 'Arial', color: '1F497D' },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({ children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LIGHT_BLUE, space: 4 } },
          children: [new TextRun({ text: 'Инструкция: заполнение географии сделок в Битрикс24', size: 18, font: 'Arial', color: '888888' })]
        })
      ]})
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: LIGHT_BLUE, space: 4 } },
          children: [
            new TextRun({ text: 'Страница ', size: 18, font: 'Arial', color: '888888' }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: 'Arial', color: '888888' }),
            new TextRun({ text: ' из ', size: 18, font: 'Arial', color: '888888' }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: 'Arial', color: '888888' }),
          ]
        })
      ]})
    },
    children: [

      // ── Заголовок ───────────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: 'Заполнение географии сделок', bold: true, size: 40, font: 'Arial', color: BLUE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: 'в Битрикс24 — инструкция для менеджеров', size: 26, font: 'Arial', color: '555555' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 320 },
        children: [new TextRun({ text: 'Май 2026', size: 20, font: 'Arial', color: '999999' })]
      }),

      divider(),

      // ── Раздел 1 ────────────────────────────────────────────────────────────
      h1('1. Зачем это нужно'),
      p('В Битрикс24 в карточке каждой сделки появилась новая вкладка «География». В ней нужно указать город, регион и федеральный округ клиента — это нужно для аналитики и отчётности по регионам.'),
      empty(80),
      p('Часть сделок уже заполнена автоматически. Остальные — в файле Excel, который вы получили.'),

      empty(160),
      divider(),

      // ── Раздел 2 ────────────────────────────────────────────────────────────
      h1('2. Файл Excel — как работать'),
      p('Вы получили файл deals_for_review.xlsx. В нём — сделки, которые нужно проверить и заполнить.'),
      empty(120),

      h2('Как отфильтровать свои сделки'),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [9026],
        rows: [new TableRow({
          children: [new TableCell({
            borders,
            width: { size: 9026, type: WidthType.DXA },
            shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 180, right: 180 },
            children: [
              li('Откройте файл deals_for_review.xlsx', true, 'nums'),
              li('В первой колонке «Ответственный» — ваше имя и фамилия', true, 'nums'),
              li('Нажмите на стрелочку автофильтра в заголовке колонки', true, 'nums'),
              li('Снимите все галочки, оставьте только своё имя → ОК', true, 'nums'),
              li('Теперь в таблице только ваши сделки', true, 'nums'),
            ]
          })]
        })]
      }),
      empty(120),

      h2('Колонки в файле'),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 6826],
        rows: [
          ['Ответственный', 'Ваше имя — используйте для фильтрации'],
          ['Стадия', 'Текущая стадия сделки'],
          ['Ссылка', 'Кликабельная ссылка — открывает сделку в Битрикс24'],
          ['Название', 'Название сделки'],
          ['Город_AI / Регион_AI / ФО_AI', 'Данные, которые определил ИИ — могут быть неточными'],
          ['Комментарий', 'Подсказка от ИИ, почему он так решил'],
        ].map(([col1, col2], i) => new TableRow({
          children: [
            new TableCell({ borders, width: { size: 2200, type: WidthType.DXA },
              shading: { fill: i % 2 === 0 ? GREY_BG : 'FFFFFF', type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 140, right: 140 },
              children: [new Paragraph({ children: [new TextRun({ text: col1, bold: true, size: 20, font: 'Arial', color: DARK_TEXT })] })] }),
            new TableCell({ borders, width: { size: 6826, type: WidthType.DXA },
              shading: { fill: i % 2 === 0 ? GREY_BG : 'FFFFFF', type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 140, right: 140 },
              children: [new Paragraph({ children: [new TextRun({ text: col2, size: 20, font: 'Arial', color: DARK_TEXT })] })] }),
          ]
        }))
      }),

      empty(160),
      divider(),

      // ── Раздел 3 ────────────────────────────────────────────────────────────
      h1('3. Как заполнить географию в сделке'),

      empty(80),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [900, 8126],
        rows: [
          stepRow(1, 'Перейдите по ссылке', 'В файле Excel кликните на ссылку в строке нужной сделки — откроется карточка сделки в Битрикс24.'),
          stepRow(2, 'Откройте вкладку «География»', 'В карточке сделки найдите вкладку «География» рядом с «Общие», «Товары» и другими вкладками. Нажмите на неё.'),
          stepRow(3, 'Выберите географию', 'Используйте виджет для указания ФО, региона и/или города (подробнее — в разделе 4).'),
          stepRow(4, 'Нажмите «Сохранить в сделку»', 'Данные запишутся в карточку сделки. Кнопка станет активной, как только вы что-то выберете.'),
        ]
      }),

      empty(160),
      divider(),

      // ── Раздел 4 ────────────────────────────────────────────────────────────
      h1('4. Как пользоваться виджетом'),

      h2('Логика автозаполнения'),
      p('Виджет умный — достаточно указать самое конкретное, что вы знаете:'),
      empty(100),
      logicTable(),
      empty(120),

      tip('Совет', 'Если знаете только город — вводите город. Регион и ФО подтянутся сами. Не нужно заполнять всё вручную.'),

      empty(160),
      h2('Выбор федерального округа (ФО)'),
      li('Нажмите на кнопку с названием ФО (ЦФО, ПФО, СФО и т.д.)'),
      li('Кнопка станет синей — ФО выбран'),
      li('Все регионы этого ФО добавятся автоматически'),
      li('Нажмите ещё раз — снимите выбор'),

      empty(120),
      h2('Поиск региона'),
      li('Начните вводить название региона в поле поиска'),
      li('Если выбран ФО — в списке будут только регионы этого ФО'),
      li('Выберите нужный из выпадающего списка'),
      li('ФО подтянется автоматически'),

      empty(120),
      h2('Поиск города'),
      li('Начните вводить название города'),
      li('Если выбран ФО или регион — список отфильтруется по ним'),
      li('Выберите город — регион и ФО подтянутся сами'),

      empty(120),
      h2('Удаление выбранных значений'),
      p('Рядом с каждым выбранным значением есть крестик ×. Нажмите его, чтобы убрать. Удаление работает только на своём уровне — удалив город, вы не потеряете регион.'),

      empty(160),
      divider(),

      // ── Раздел 5 ────────────────────────────────────────────────────────────
      h1('5. Что ставить, если не знаете точный город'),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3000, 6026],
        rows: [
          new TableRow({ tableHeader: true, children: [
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA },
              shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 140, right: 140 },
              children: [new Paragraph({ children: [new TextRun({ text: 'Ситуация', bold: true, size: 21, font: 'Arial', color: 'FFFFFF' })] })] }),
            new TableCell({ borders, width: { size: 6026, type: WidthType.DXA },
              shading: { fill: BLUE, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 140, right: 140 },
              children: [new Paragraph({ children: [new TextRun({ text: 'Что делать', bold: true, size: 21, font: 'Arial', color: 'FFFFFF' })] })] }),
          ]}),
          ...([
            ['Знаете точный город', 'Введите город — этого достаточно'],
            ['Знаете только регион', 'Выберите регион через поиск'],
            ['Знаете только ФО', 'Нажмите кнопку ФО'],
            ['Клиент работает по всей России', 'Выберите все 8 ФО'],
            ['Несколько городов/регионов', 'Добавляйте их по одному — мультивыбор поддерживается'],
            ['Совсем не знаете географию', 'Оставьте пустым и перейдите к следующей сделке'],
          ].map(([col1, col2], i) => new TableRow({ children: [
            new TableCell({ borders, width: { size: 3000, type: WidthType.DXA },
              shading: { fill: i % 2 === 0 ? GREY_BG : 'FFFFFF', type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 140, right: 140 },
              children: [new Paragraph({ children: [new TextRun({ text: col1, bold: true, size: 21, font: 'Arial', color: DARK_TEXT })] })] }),
            new TableCell({ borders, width: { size: 6026, type: WidthType.DXA },
              shading: { fill: i % 2 === 0 ? GREY_BG : 'FFFFFF', type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 140, right: 140 },
              children: [new Paragraph({ children: [new TextRun({ text: col2, size: 21, font: 'Arial', color: DARK_TEXT })] })] }),
          ]}))),
        ]
      }),

      empty(160),
      divider(),

      // ── Раздел 6 ────────────────────────────────────────────────────────────
      h1('6. Частые вопросы'),

      p('Кнопка «Сохранить» неактивна?', { bold: true }),
      p('Нужно выбрать хотя бы одно значение (ФО, регион или город).'),
      empty(80),

      p('Виджет показывает «Загрузка»?', { bold: true }),
      p('Обновите страницу сделки (F5 или Cmd+R).'),
      empty(80),

      p('Нет нужного города в списке?', { bold: true }),
      p('Выберите регион — этого достаточно для аналитики.'),
      empty(80),

      p('Сохранил неправильно — можно исправить?', { bold: true }),
      p('Да. Снова откройте вкладку «География», внесите изменения и нажмите «Сохранить в сделку».'),

      empty(160),
      tip('Вопросы и проблемы', 'Если виджет не работает или что-то непонятно — напишите Тойгильдиной Марии.'),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/Users/marusia/projects/work/bitrix24-geo-widget/Инструкция_география_Битрикс24.docx', buffer);
  console.log('Готово!');
});
