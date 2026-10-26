/**
 * Research Reference Tracker - one-click setup
 * Run setupAll() ONCE. It creates the Google Form, a linked Google Sheet,
 * a trigger that fills in DOI details and flags duplicates, and an
 * optional weekly digest email for the professor.
 */

const CONFIG = {
  FORM_TITLE: 'EDUC 6181 - Add a Research Reference',
  SHEET_NAME: 'EDUC 6181 Reference Database',
  PROF_EMAIL: 'mduv@gse.upenn.edu',            
  TOPICS: [
    'AI in Education',
    'STEM Integration / Teacher Learning',
    'Collaborative Learning / SSRL',
    'Instructional Video',
    'Parenting & Motivation',
    'Research Methods',
    'Other'
  ],
  KEYWORDS: [
    'Generative AI', 'AI literacy', 'Critical thinking', 'Collaboration',
    'Self-regulation', 'Motivation', 'Teacher education', 'Video analysis',
    'Learning analytics', 'Parent involvement', 'Deepfakes / media literacy'
  ],
  METHODS: [
    'Systematic review', 'Survey / questionnaire', 'Experiment',
    'Qualitative', 'Mixed methods', 'Theoretical / conceptual', 'Other'
  ],
  AGE_GROUPS: [
    'Early childhood', 'Elementary', 'Middle school', 'High school',
    'Higher education', 'Teachers / educators', 'Not specified'
  ]
};

const EXTRA_HEADERS = ['Title', 'Authors', 'Year', 'Journal', 'APA Citation', 'Status'];

/** RUN THIS ONCE. */
function setupAll() {
  const form = FormApp.create(CONFIG.FORM_TITLE);
  form.setDescription(
    'Add one paper per submission. Just paste the DOI - the title, authors, year, ' +
    'journal and APA citation are filled in automatically.'
  );
  form.setCollectEmail(true);
  form.setLimitOneResponsePerUser(false);
  form.setPublishingSummary(true);
  form.setShowLinkToRespondAgain(true);

  form.addTextItem()
    .setTitle('DOI')
    .setHelpText('Example: 10.1016/j.caeai.2021.100040 (a full https://doi.org/... link also works)')
    .setRequired(true);

  form.addListItem()
    .setTitle('Primary topic')
    .setChoiceValues(CONFIG.TOPICS)
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('Keywords (pick all that apply)')
    .setChoiceValues(CONFIG.KEYWORDS)
    .showOtherOption(true);

  form.addListItem()
    .setTitle('Method')
    .setChoiceValues(CONFIG.METHODS)
    .setRequired(true);

  form.addListItem()
    .setTitle('Age group / setting')
    .setChoiceValues(CONFIG.AGE_GROUPS)
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('Why is this relevant to our research?')
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('Useful?')
    .setChoiceValues(['Yes, this was useful']);

  const ss = SpreadsheetApp.create(CONFIG.SHEET_NAME);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  if (CONFIG.PROF_EMAIL) {
    ScriptApp.newTrigger('weeklyDigest')
      .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  }

  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

  Logger.log('FORM (share this with students): ' + form.getPublishedUrl());
  Logger.log('FORM (edit): ' + form.getEditUrl());
  Logger.log('SHEET (professor): ' + ss.getUrl());
}

/** Runs automatically on every form submission. */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const cols = ensureColumns_(sheet);

  const rawDoi = String(sheet.getRange(row, cols['DOI']).getValue());
  const doi = normalizeDoi_(rawDoi);
  sheet.getRange(row, cols['DOI']).setValue(doi);

  const last = sheet.getLastRow();
  if (last > 2) {
    const dois = sheet.getRange(2, cols['DOI'], last - 1, 1).getValues();
    for (let i = 0; i < dois.length; i++) {
      const r = i + 2;
      if (r !== row && normalizeDoi_(String(dois[i][0])) === doi) {
        sheet.getRange(row, cols['Status']).setValue('DUPLICATE of row ' + r);
        return;
      }
    }
  }

  const meta = lookupCrossref_(doi);
  if (!meta) {
    sheet.getRange(row, cols['Status']).setValue('DOI not found - check manually');
    return;
  }
  sheet.getRange(row, cols['Title']).setValue(meta.title);
  sheet.getRange(row, cols['Authors']).setValue(meta.authors);
  sheet.getRange(row, cols['Year']).setValue(meta.year);
  sheet.getRange(row, cols['Journal']).setValue(meta.journal);
  sheet.getRange(row, cols['APA Citation']).setValue(meta.apa);
  sheet.getRange(row, cols['Status']).setValue('OK');
}

/** Weekly summary email to the professor. */
function weeklyDigest() {
  if (!CONFIG.PROF_EMAIL) return;
  const ss = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty('SS_ID'));
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const iTime = h.indexOf('Timestamp');
  const iTopic = h.indexOf('Primary topic');
  const iTitle = h.indexOf('Title');
  const iWho = h.indexOf('Email Address');
  const cutoff = new Date(Date.now() - 7 * 24 * 3600 * 1000);

  const recent = data.slice(1).filter(r => r[iTime] instanceof Date && r[iTime] >= cutoff);
  const byTopic = {};
  recent.forEach(r => { byTopic[r[iTopic]] = (byTopic[r[iTopic]] || 0) + 1; });

  let body = recent.length + ' new reference(s) this week.\n\nBy topic:\n';
  Object.keys(byTopic).forEach(t => { body += '  - ' + t + ': ' + byTopic[t] + '\n'; });
  body += '\nNew entries:\n';
  recent.forEach(r => { body += '  - ' + (r[iTitle] || '(title pending)') + ' [' + r[iWho] + ']\n'; });
  body += '\nSheet: ' + ss.getUrl();

  MailApp.sendEmail(CONFIG.PROF_EMAIL, 'Weekly research reference digest', body);
}

/* ------------------------- helpers ------------------------- */

function ensureColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  EXTRA_HEADERS.forEach(name => {
    if (headers.indexOf(name) === -1) {
      sheet.getRange(1, headers.length + 1).setValue(name);
      headers.push(name);
    }
  });

  const map = {};
  headers.forEach((name, i) => { map[name] = i + 1; });
  return map;
}

function normalizeDoi_(s) {
  return s.trim().toLowerCase()
    .replace(/^https?:\/\/(dx\.)?doi\.org\//, '')
    .replace(/^doi:\s*/, '')
    .trim();
}

function lookupCrossref_(doi) {
  try {
    const url = 'https://api.crossref.org/works/' + encodeURIComponent(doi);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    const m = JSON.parse(res.getContentText()).message;

    const title = (m.title && m.title[0]) || '';
    const journal = (m['container-title'] && m['container-title'][0]) || '';
    const issued = (m.issued && m.issued['date-parts'] && m.issued['date-parts'][0]) || [];
    const year = issued[0] || '';
    const authorList = (m.author || []).map(a => ({ family: a.family || a.name || '', given: a.given || '' }));

    const authorsPlain = authorList.map(a => (a.given + ' ' + a.family).trim()).join(', ');
    const apaAuthors = formatApaAuthors_(authorList);

    let apa = apaAuthors + ' (' + year + '). ' + title + '. ' + journal;
    if (m.volume) apa += ', ' + m.volume;
    if (m.issue) apa += '(' + m.issue + ')';
    if (m.page) apa += ', ' + m.page;
    apa += '. https://doi.org/' + doi;

    return { title: title, authors: authorsPlain, year: year, journal: journal, apa: apa };
  } catch (err) {
    return null;
  }
}

function formatApaAuthors_(list) {
  const fmt = a => {
    const initials = a.given.split(/[\s-]+/).filter(Boolean)
      .map(p => p.charAt(0).toUpperCase() + '.').join(' ');
    return (a.family + (initials ? ', ' + initials : '')).trim();
  };
  const names = list.map(fmt);
  if (names.length === 0) return '';
  if (names.length === 1) return names[0];
  if (names.length <= 20) return names.slice(0, -1).join(', ') + ', & ' + names[names.length - 1];
  return names.slice(0, 19).join(', ') + ', ... ' + names[names.length - 1];
}
