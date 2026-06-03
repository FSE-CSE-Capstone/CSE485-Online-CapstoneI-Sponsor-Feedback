// CSE 485 Online Capstone I Sponsor Feedback Survey – scripts.js
// Complete rewrite: bug fixes + semester isolation + validation + report generation
(function () {
  'use strict';

  // ─────────────────────────────────────────────────────────────────────────────
  // SEMESTER-AWARE STORAGE KEY
  // window.SURVEY_ROUND is set in index.html – change it each semester.
  // Any saved progress from a different round is automatically removed.
  // ─────────────────────────────────────────────────────────────────────────────
  var BASE_KEY    = 'sponsor_progress_v1';
  var ROUND       = (window.SURVEY_ROUND || 'round1');
  var STORAGE_KEY = BASE_KEY + '_' + ROUND;

  try {
    Object.keys(localStorage).forEach(function (k) {
      if (k.indexOf(BASE_KEY + '_') === 0 && k !== STORAGE_KEY) {
        localStorage.removeItem(k);
      }
    });
  } catch (e) { console.warn('Round cleanup failed', e); }

  // ─────────────────────────────────────────────────────────────────────────────
  // CONFIG – CSE 485 Online Cloudflare Workers
  // ─────────────────────────────────────────────────────────────────────────────
  var ENDPOINT_URL    = 'https://cse485-online-worker.sbecerr7.workers.dev/';
  var DATA_LOADER_URL = 'https://cse485-online-data-loader.sbecerr7.workers.dev/';

   // ─────────────────────────────────────────────────────────────────────────────
  // RUBRIC
  // ─────────────────────────────────────────────────────────────────────────────
  var RUBRIC = [
    {
      title: 'Development Effort',
      description: 'Student has contributed an appropriate amount of development effort towards this project. Development effort should be balanced between all team members; student should commit to a fair amount of development effort on each sprint.'
    },
    {
      title: 'Meetings',
      description: 'Students are expected to be proactive. Contributions and participation in meetings help ensure the student is aware of project goals.'
    },
    {
      title: 'Understanding',
      description: 'Students are expected to understand important details of the project and be able to explain it from different stakeholder perspectives.'
    },
    {
      title: 'Quality',
      description: 'Students should complete assigned work to a high quality: correct, documented, and self-explanatory where appropriate.'
    },
    {
      title: 'Communication',
      description: 'Students are expected to be in regular communication and maintain professionalism when interacting with the sponsor.'
    }
  ];

  // ─────────────────────────────────────────────────────────────────────────────
  // DOM REFERENCES
  // ─────────────────────────────────────────────────────────────────────────────
  var $ = function (id) { return document.getElementById(id); };

  var stageIdentity    = $('stage-identity');
  var stageProjects    = $('stage-projects');
  var stageThankyou    = $('stage-thankyou');
  var identitySubmit   = $('identitySubmit');
  var backToIdentity   = $('backToIdentity');
  var nameInput        = $('fullName');
  var emailInput       = $('email');
  var projectListEl    = $('project-list');
  var matrixContainer  = $('matrix-container');
  var formStatus       = $('form-status');
  var submitProjectBtn = $('submitProject');
  var finishStartOverBtn = $('finishStartOver');
  var downloadReportBtn  = $('downloadReport');
  var printReportBtn     = $('printReport');
  var welcomeBlock     = $('welcome-block');
  var underTitle       = $('under-title');
  var progressCounter  = $('progress-counter');

  // ─────────────────────────────────────────────────────────────────────────────
  // STATE
  // ─────────────────────────────────────────────────────────────────────────────
  var sponsorData        = {};   // full map: email → { projects: { name: [students] } }
  var sponsorProjects    = {};   // current user's projects (populated after login)
  var currentEmail       = '';
  var currentName        = '';
  var currentProject     = '';
  var completedProjects  = {};   // { projectName: true }
  var remoteCompletedProjects = {}; // submitted by any sponsor, loaded from the worker
  var stagedRatings      = {};   // in-progress draft ratings
  var submittedResponses = {};   // full payloads of submitted projects (for report)

  // ─────────────────────────────────────────────────────────────────────────────
  // UTILITY HELPERS
  // ─────────────────────────────────────────────────────────────────────────────

  function setStatus(msg, type) {
    if (!formStatus) return;
    formStatus.textContent = msg || '';
    formStatus.className = 'form-status' + (type ? ' form-status-' + type : '');
  }

  function escapeHtml(s) {
    if (s == null) return '';
    return String(s).replace(/[&<>"']/g, function (m) {
      return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[m];
    });
  }

  // Safe CSS.escape polyfill
  function cssEscape(str) {
    if (window.CSS && CSS.escape) return CSS.escape(str);
    return str.replace(/([^\w-])/g, '\\$1');
  }

  // Tiny DOM builder
  function el(tag, props, children) {
    var n = document.createElement(tag);
    if (props) {
      Object.keys(props).forEach(function (k) {
        if      (k === 'class') n.className = props[k];
        else if (k === 'html')  n.innerHTML = props[k];
        else if (k === 'text')  n.textContent = props[k];
        else if (k === 'style') Object.assign(n.style, props[k]);
        else n.setAttribute(k, props[k]);
      });
    }
    if (children) {
      children.forEach(function (c) {
        if (typeof c === 'string') n.appendChild(document.createTextNode(c));
        else n.appendChild(c);
      });
    }
    return n;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // BUILD SPONSOR MAP FROM GOOGLE SHEETS ROWS
  // Tolerant of varied column naming, multiple emails per cell, etc.
  // ─────────────────────────────────────────────────────────────────────────────
  function buildSponsorMap(rows) {
    var map = {};
    if (!Array.isArray(rows)) return map;

    var emailRegex = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;

    function cleanToken(tok) {
      if (!tok) return '';
      return tok.replace(/^[\s"'`([{]+|[\s"'`)\]}.,:;]+$/g, '').replace(/\u00A0/g, ' ').trim();
    }

    rows.forEach(function (rawRow) {
      var project = '', student = '', sponsorCell = '';

      Object.keys(rawRow || {}).forEach(function (rawKey) {
        var keyNorm = String(rawKey || '').trim().toLowerCase();
        var rawVal  = (rawRow[rawKey] || '').toString().replace(/\u00A0/g, ' ').trim();

        if (!project && /^(project|project name|project_title|group_name|projectname)$/.test(keyNorm))
          project = rawVal;
        else if (!student && /^(student|student name|students|name|student_name)$/.test(keyNorm))
          student = rawVal;
        else if (!sponsorCell && /^(sponsoremail|sponsor email|sponsor|email|login_id|sponsor_email)$/.test(keyNorm))
          sponsorCell = rawVal;
      });

      // Fallback: extract any email from any cell
      if (!sponsorCell) {
        var fallback = [];
        Object.keys(rawRow || {}).forEach(function (k) {
          var found = (rawRow[k] || '').toString().match(emailRegex);
          if (found) fallback = fallback.concat(found);
        });
        if (fallback.length) sponsorCell = fallback.join(', ');
      }

      project = (project || '').trim();
      student = (student || '').trim();
      if (!sponsorCell || !project || !student) return;

      var foundEmails = [];
      sponsorCell.split(/[,;\/|]+/).forEach(function (t) {
        var cleaned = cleanToken(t);
        if (!cleaned) return;
        var matches = cleaned.match(emailRegex) || t.match(emailRegex) || [];
        matches.forEach(function (em) { foundEmails.push(em.toLowerCase().trim()); });
      });

      var unique = [];
      foundEmails.forEach(function (e) {
        if (!e || e.indexOf('@') === -1) return;
        var parts = e.split('@');
        if (parts.length !== 2 || parts[1].indexOf('.') === -1) return;
        if (unique.indexOf(e) === -1) unique.push(e);
      });

      if (!unique.length) return;
      unique.forEach(function (email) {
        if (!map[email]) map[email] = { projects: {} };
        if (!map[email].projects[project]) map[email].projects[project] = [];
        if (map[email].projects[project].indexOf(student) === -1)
          map[email].projects[project].push(student);
      });
    });

    return map;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // PERSISTENCE (localStorage)
  // ─────────────────────────────────────────────────────────────────────────────
  function saveProgress() {
    var payload = {
      name:               currentName,
      email:              currentEmail,
      completedProjects:  completedProjects,
      stagedRatings:      stagedRatings,
      submittedResponses: submittedResponses
    };
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(payload)); }
    catch (e) { console.warn('Could not save progress', e); }
  }

  // Returns true if usable saved state was found
  function loadProgress() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return false;
      var obj = JSON.parse(raw);
      if (obj && obj.email) {
        currentName        = obj.name || '';
        currentEmail       = obj.email || '';
        completedProjects  = obj.completedProjects  || {};
        stagedRatings      = obj.stagedRatings      || {};
        submittedResponses = obj.submittedResponses  || {};
        if (nameInput)  nameInput.value  = currentName;
        if (emailInput) emailInput.value = currentEmail;
        return true;
      }
    } catch (e) { console.warn('Could not load progress', e); }
    return false;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // PROGRESS COUNTER  (e.g. "2 of 3 projects completed")
  // ─────────────────────────────────────────────────────────────────────────────
  function updateProgressCounter() {
    if (!progressCounter) return;
    var entry = sponsorData[currentEmail];
    if (!entry || !entry.projects) { progressCounter.textContent = ''; return; }
    var projects = Object.keys(entry.projects);
    var total = projects.length;
    var done  = projects.filter(function (p) { return isProjectCompleted(p); }).length;
    progressCounter.textContent = done + ' of ' + total + ' project' + (total !== 1 ? 's' : '') + ' completed';
  }

  function isProjectCompleted(projectName) {
    return !!completedProjects[projectName] || !!remoteCompletedProjects[projectName];
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // POPULATE PROJECT LIST
  // ─────────────────────────────────────────────────────────────────────────────
  function populateProjectListFor(email) {
    if (!projectListEl) return;
    projectListEl.innerHTML = '';
    sponsorProjects = {};

    var entry = sponsorData[email];
    if (!entry || !entry.projects) {
      setStatus('No projects found for that email address.', 'error');
      return;
    }

    var allProjects = Object.keys(entry.projects).slice();
    // Sort: incomplete projects first
    allProjects.sort(function (a, b) {
      return (isProjectCompleted(a) ? 1 : 0) - (isProjectCompleted(b) ? 1 : 0);
    });

    allProjects.forEach(function (p) {
      var isDone = isProjectCompleted(p);
      var isRemoteOnly = !completedProjects[p] && !!remoteCompletedProjects[p];

      var li = el('li', {
        class:       'project-item' + (isDone ? ' completed' : ''),
        tabindex:    isDone ? '-1' : '0',
        'data-project': p,
        'aria-label': p + (isDone ? ' (already submitted)' : '')
      });

      var nameSpan = el('span', { class: 'project-item-name', text: p });
      li.appendChild(nameSpan);

      if (isDone) {
        li.appendChild(el('span', { class: 'meta', text: isRemoteOnly ? '\u2713 Already submitted' : '\u2713 Completed' }));
      } else {
        li.appendChild(el('span', { class: 'project-item-arrow', text: '\u2192' }));
      }

      if (!isDone) {
        li.addEventListener('click', function () {
          // Deactivate previous selection
          Array.from(projectListEl.querySelectorAll('.project-item.active')).forEach(function (a) {
            a.classList.remove('active');
          });
          li.classList.add('active');
          currentProject = p;
          loadProjectIntoMatrix(p, entry.projects[p]);
          setStatus('');
          // Scroll to matrix after a short delay for render
          setTimeout(function () {
            var info = $('matrix-info');
            if (info) info.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }, 120);
        });

        li.addEventListener('keydown', function (e) {
          if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); li.click(); }
        });
      }

      projectListEl.appendChild(li);
      sponsorProjects[p] = entry.projects[p].slice();
    });

    updateProgressCounter();
    setStatus('');
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // VALIDATION – ensure each criterion has at least one rating
  // (student OR team-overall counts for that criterion)
  // ─────────────────────────────────────────────────────────────────────────────
  function getRatingMode(students) {
    var hasStudentRatings = false;
    var hasTeamRatings = false;

    for (var c = 0; c < RUBRIC.length; c++) {
      if (document.querySelector('input[name="rating-' + c + '-team"]:checked')) {
        hasTeamRatings = true;
      }
      for (var s = 0; s < students.length; s++) {
        if (document.querySelector('input[name="rating-' + c + '-' + s + '"]:checked')) {
          hasStudentRatings = true;
        }
      }
    }

    if (hasStudentRatings && hasTeamRatings) return { mode: 'mixed' };
    if (hasTeamRatings) return { mode: 'team' };
    if (hasStudentRatings) return { mode: 'individual' };
    return { mode: 'empty' };
  }

  function validateRatings(students, ratingMode) {
    var issues = [];
    var mode = (ratingMode && ratingMode.mode) || getRatingMode(students).mode;

    if (mode === 'empty') {
      issues.push('Choose either individual student ratings or the Team Overall row.');
      return issues;
    }

    if (mode === 'mixed') {
      issues.push('Choose one rating method: individual students or Team Overall, not both.');
      return issues;
    }

    for (var c = 0; c < RUBRIC.length; c++) {
      if (mode === 'team') {
        if (!document.querySelector('input[name="rating-' + c + '-team"]:checked')) {
          issues.push('Missing Team Overall score for "' + RUBRIC[c].title + '".');
          return issues;
        }
      } else {
        for (var s = 0; s < students.length; s++) {
          if (!document.querySelector('input[name="rating-' + c + '-' + s + '"]:checked')) {
            issues.push('Missing score for ' + students[s] + ' on "' + RUBRIC[c].title + '".');
            return issues;
          }
        }
      }
    }
    return issues;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // BUILD MATRIX FOR SELECTED PROJECT
  // ─────────────────────────────────────────────────────────────────────────────
  function loadProjectIntoMatrix(projectName, students) {
    if (!projectName) return;
    currentProject = projectName;

    // Remove any previously rendered matrix/comment UI
    var existingInfo = $('matrix-info');
    if (existingInfo && existingInfo.parentNode) existingInfo.parentNode.removeChild(existingInfo);

    var oldComment = document.querySelector('.section.section-comment');
    if (oldComment && oldComment.parentNode) oldComment.parentNode.removeChild(oldComment);

    // Build the matrix-info header block
    var info = el('div', { id: 'matrix-info', class: 'matrix-info-block' });
    info.appendChild(el('div', {
      class: 'current-project-header',
      text: 'Evaluating: ' + projectName
    }));
    info.appendChild(el('div', {
      class: 'matrix-info-desc',
      text: 'Rate either every student (1\u20137) or only the Team Overall row. Do not use both methods for the same project.'
    }));

    if (matrixContainer && matrixContainer.parentNode) {
      matrixContainer.parentNode.insertBefore(info, matrixContainer);
    }

    if (!students || !students.length) {
      if (matrixContainer) matrixContainer.textContent = 'No students found for this project.';
      return;
    }

    if (!stagedRatings[currentProject]) stagedRatings[currentProject] = {};

    var temp = document.createElement('div');

    RUBRIC.forEach(function (crit, cIdx) {
      var card = el('div', { class: 'card matrix-card' });

      // Criterion header
      var critWrap = el('div', { class: 'matrix-criterion' });
      critWrap.appendChild(el('h4', {
        class: 'matrix-criterion-title',
        text: (cIdx + 1) + '. ' + crit.title
      }));
      critWrap.appendChild(el('div', {
        class: 'matrix-criterion-desc',
        text: crit.description
      }));

      // Scrollable table wrapper (for mobile)
      var tableWrap = el('div', { class: 'table-scroll-wrap' });
      var table = el('table', { class: 'matrix-table', role: 'grid' });

      // Colgroup for proportional widths
      var colgroup = el('colgroup');
      colgroup.appendChild(el('col', { class: 'col-student-def' }));
      colgroup.appendChild(el('col', { class: 'col-desc-def' }));
      for (var ci = 0; ci < 7; ci++) colgroup.appendChild(el('col', { class: 'col-radio-def' }));
      colgroup.appendChild(el('col', { class: 'col-desc-def' }));
      table.appendChild(colgroup);

      // Header row
      var thead = el('thead');
      var trHead = el('tr');
      trHead.appendChild(el('th', { scope: 'col', class: 'col-student', text: 'Student' }));
      trHead.appendChild(el('th', { scope: 'col', class: 'header-descriptor',
        html: '<div class="hd-line">Far Below</div><div class="hd-sub">Expectations</div>' }));
      for (var k = 1; k <= 7; k++) {
        trHead.appendChild(el('th', { scope: 'col', class: 'col-score-num', text: String(k) }));
      }
      trHead.appendChild(el('th', { scope: 'col', class: 'header-descriptor header-descriptor-right',
        html: '<div class="hd-line">Exceeds</div><div class="hd-sub">Expectations</div>' }));
      thead.appendChild(trHead);
      table.appendChild(thead);

      var tbody = el('tbody');

      // One row per student
      students.forEach(function (studentName, sIdx) {
        var tr = el('tr', { class: sIdx % 2 === 0 ? 'row-even' : 'row-odd' });
        tr.appendChild(el('td', { class: 'col-student', text: studentName }));
        tr.appendChild(el('td', { class: 'col-descriptor' }));

        for (var score = 1; score <= 7; score++) {
          var inputId = 'r-' + cIdx + '-' + sIdx + '-' + score;
          var inp = el('input', {
            type:  'radio',
            name:  'rating-' + cIdx + '-' + sIdx,
            value: String(score),
            id:    inputId,
            'aria-label': studentName + ', score ' + score
          });
          var staged = (stagedRatings[currentProject][sIdx] || {})[cIdx];
          if (staged && String(staged) === String(score)) inp.checked = true;

          var lbl = el('label', { for: inputId, class: 'radio-label' });
          lbl.appendChild(inp);
          var td = el('td', { class: 'col-radio' });
          td.appendChild(lbl);
          tr.appendChild(td);
        }
        tr.appendChild(el('td', { class: 'col-descriptor' }));
        tbody.appendChild(tr);
      });

      // Team Overall row
      var trTeam = el('tr', { class: 'row-team' });
      trTeam.appendChild(el('td', { class: 'col-student',
        html: '<strong>Team Overall</strong>' }));
      trTeam.appendChild(el('td', { class: 'col-descriptor' }));
      for (var ts = 1; ts <= 7; ts++) {
        var tInputId = 'r-' + cIdx + '-team-' + ts;
        var tInp = el('input', {
          type:  'radio',
          name:  'rating-' + cIdx + '-team',
          value: String(ts),
          id:    tInputId,
          'aria-label': 'Team overall, score ' + ts
        });
        var stagedTeam = (stagedRatings[currentProject].team || {})[cIdx];
        if (stagedTeam && String(stagedTeam) === String(ts)) tInp.checked = true;

        var tLbl = el('label', { for: tInputId, class: 'radio-label' });
        tLbl.appendChild(tInp);
        var tTd = el('td', { class: 'col-radio' });
        tTd.appendChild(tLbl);
        trTeam.appendChild(tTd);
      }
      trTeam.appendChild(el('td', { class: 'col-descriptor' }));
      tbody.appendChild(trTeam);

      table.appendChild(tbody);
      tableWrap.appendChild(table);
      critWrap.appendChild(tableWrap);
      card.appendChild(critWrap);
      temp.appendChild(card);
    });

    // Replace matrix container contents
    if (matrixContainer) {
      matrixContainer.innerHTML = '';
      while (temp.firstChild) matrixContainer.appendChild(temp.firstChild);
    }

    renderCommentSection(projectName, students);
    attachMatrixListeners();
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // COMMENT SECTION (per-student + team)
  // ─────────────────────────────────────────────────────────────────────────────
  function renderCommentSection(projectName, students) {
    var old = document.querySelector('.section.section-comment');
    if (old && old.parentNode) old.parentNode.removeChild(old);

    var sec = el('div', { class: 'section section-comment' });
    sec.appendChild(el('h3', { class: 'section-title', text: 'Additional Comments (Optional)' }));

    var stagedComments = (stagedRatings[projectName] || {})._studentComments || {};

    students.forEach(function (studentName, sIdx) {
      var panel = buildCommentPanel(
        studentName,
        'comment-public-' + sIdx,
        'comment-private-' + sIdx,
        (stagedComments[studentName] || {}).public,
        (stagedComments[studentName] || {}).private
      );
      sec.appendChild(panel);
    });

    // Team overall comments
    var stagedGroup = (stagedRatings[projectName] || {})._groupComments || {};
    var groupPanel = buildCommentPanel(
      'Team Overall',
      'comment-group-public',
      'comment-group-private',
      stagedGroup.public,
      stagedGroup.private
    );
    sec.appendChild(groupPanel);

    if (matrixContainer && matrixContainer.parentNode) {
      if (matrixContainer.nextSibling) {
        matrixContainer.parentNode.insertBefore(sec, matrixContainer.nextSibling);
      } else {
        matrixContainer.parentNode.appendChild(sec);
      }
    }

    // Attach textarea auto-save listeners
    Array.from(sec.querySelectorAll('textarea')).forEach(function (ta) {
      ta.addEventListener('input', saveDraftHandler);
    });
  }

  function buildCommentPanel(label, pubId, privId, pubVal, privVal) {
    var wrapper = el('div', { class: 'student-comment-panel' });
    var headerRow = el('div', { class: 'comment-panel-header' });
    headerRow.appendChild(el('span', { class: 'comment-panel-name', text: label }));

    var toggleBtn = el('button', {
      type:  'button',
      class: 'btn btn-mini comment-toggle',
      text:  '\u25be Add comment'
    });
    headerRow.appendChild(toggleBtn);
    wrapper.appendChild(headerRow);

    var content = el('div', { class: 'student-comment-content' });
    content.style.display = 'none';

    content.appendChild(el('div', { class: 'comment-label', text: 'Comments to be SHARED WITH THE STUDENT' }));
    var taPublic = el('textarea', {
      id:          pubId,
      placeholder: 'Comments to share with student\u2026',
      rows:        '3'
    });
    content.appendChild(taPublic);

    content.appendChild(el('div', { class: 'comment-label', text: 'Comments to be SHARED ONLY WITH THE INSTRUCTOR' }));
    var taPrivate = el('textarea', {
      id:          privId,
      placeholder: 'Private comments for instructor\u2026',
      rows:        '3'
    });
    content.appendChild(taPrivate);

    // Restore staged values
    if (pubVal)  taPublic.value  = pubVal;
    if (privVal) taPrivate.value = privVal;
    if ((pubVal && pubVal.length) || (privVal && privVal.length)) {
      content.style.display = 'block';
      toggleBtn.textContent = '\u25b4 Hide comment';
    }

    toggleBtn.addEventListener('click', function () {
      if (content.style.display === 'none') {
        content.style.display = 'block';
        toggleBtn.textContent = '\u25b4 Hide comment';
      } else {
        content.style.display = 'none';
        toggleBtn.textContent = '\u25be Add comment';
      }
    });

    wrapper.appendChild(content);
    return wrapper;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // AUTO-SAVE DRAFT
  // ─────────────────────────────────────────────────────────────────────────────
  function attachMatrixListeners() {
    if (!matrixContainer) return;
    matrixContainer.removeEventListener('change', saveDraftHandler);
    matrixContainer.removeEventListener('input',  saveDraftHandler);
    matrixContainer.addEventListener('change', saveDraftHandler);
    matrixContainer.addEventListener('input',  saveDraftHandler);
  }

  function saveDraftHandler() {
    if (!currentProject) return;
    var students = sponsorProjects[currentProject] || [];
    var draft = stagedRatings[currentProject] || {};

    for (var s = 0; s < students.length; s++) {
      draft[s] = draft[s] || {};
      for (var c = 0; c < RUBRIC.length; c++) {
        var sel = document.querySelector('input[name="rating-' + c + '-' + s + '"]:checked');
        draft[s][c] = sel ? parseInt(sel.value, 10) : null;
      }
    }

    draft.team = draft.team || {};
    for (var ct = 0; ct < RUBRIC.length; ct++) {
      var selT = document.querySelector('input[name="rating-' + ct + '-team"]:checked');
      draft.team[ct] = selT ? parseInt(selT.value, 10) : null;
    }

    draft._studentComments = draft._studentComments || {};
    for (var i = 0; i < students.length; i++) {
      var sName   = students[i];
      var pubEl   = document.getElementById('comment-public-' + i);
      var privEl  = document.getElementById('comment-private-' + i);
      draft._studentComments[sName] = {
        public:  (pubEl  && pubEl.value)  || '',
        private: (privEl && privEl.value) || ''
      };
    }

    var gpPub  = document.getElementById('comment-group-public');
    var gpPriv = document.getElementById('comment-group-private');
    draft._groupComments = {
      public:  (gpPub  && gpPub.value)  || '',
      private: (gpPriv && gpPriv.value) || ''
    };

    stagedRatings[currentProject] = draft;
    saveProgress();
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // SUBMIT CURRENT PROJECT
  // ─────────────────────────────────────────────────────────────────────────────
  function submitCurrentProject() {
    if (!currentProject) {
      setStatus('No project selected. Click a project from the list above.', 'error');
      return;
    }

    var students = sponsorProjects[currentProject] || [];
    if (!students.length) { setStatus('No students found for this project.', 'error'); return; }

    // Snapshot draft before validation
    saveDraftHandler();

    // Validate completeness and rating mode
    var ratingMode = getRatingMode(students);
    var issues = validateRatings(students, ratingMode);
    if (issues.length) {
      setStatus('Please complete the evaluation before submitting. ' + issues[0], 'error');
      return;
    }

    var responses = [];

    if (ratingMode.mode === 'individual') {
      for (var s = 0; s < students.length; s++) {
        var ratingsObj = {};
        for (var c = 0; c < RUBRIC.length; c++) {
          var sel = document.querySelector('input[name="rating-' + c + '-' + s + '"]:checked');
          ratingsObj[RUBRIC[c].title] = sel ? parseInt(sel.value, 10) : null;
        }
        responses.push({
          student:           students[s],
          ratings:           ratingsObj,
          commentShared:     (document.getElementById('comment-public-'  + s) || {}).value || '',
          commentInstructor: (document.getElementById('comment-private-' + s) || {}).value || '',
          isTeam:            false
        });
      }
    } else if (ratingMode.mode === 'team') {
      var teamRatingsObj = {};
      for (var tc = 0; tc < RUBRIC.length; tc++) {
        var tSel = document.querySelector('input[name="rating-' + tc + '-team"]:checked');
        teamRatingsObj[RUBRIC[tc].title] = tSel ? parseInt(tSel.value, 10) : null;
      }
      var gpPub  = (document.getElementById('comment-group-public')  || {}).value || '';
      var gpPriv = (document.getElementById('comment-group-private') || {}).value || '';
      responses.push({
        student:           'Team Overall',
        ratings:           teamRatingsObj,
        commentShared:     gpPub,
        commentInstructor: gpPriv,
        isTeam:            true
      });
    }

    var submittingProject = currentProject;
    var payload = {
      sponsorName:  currentName,
      sponsorEmail: currentEmail,
      project:      submittingProject,
      surveyRound:  ROUND,
      rubric:       RUBRIC.map(function (r) { return r.title; }),
      responses:    responses,
      timestamp:    new Date().toISOString()
    };

    // Cache payload for report generation before the async call
    submittedResponses[submittingProject] = payload;
    saveProgress();

    setStatus('Submitting\u2026', 'info');
    if (submitProjectBtn) submitProjectBtn.disabled = true;

    fetch(ENDPOINT_URL, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(payload)
    })
    .then(function (resp) {
      if (!resp.ok) {
        return resp.text().then(function (txt) {
          var parsed = null;
          try { parsed = JSON.parse(txt); } catch (_) {}
          var err = new Error('Server error ' + resp.status + ': ' + txt);
          if (parsed && parsed.alreadySubmitted) {
            err.alreadySubmitted = true;
            err.project = parsed.project || submittingProject;
          }
          throw err;
        });
      }
      return resp.json().catch(function () { return {}; });
    })
    .then(function () {
      completedProjects[submittingProject] = true;
      remoteCompletedProjects[submittingProject] = true;
      delete stagedRatings[submittingProject];
      saveProgress();

      populateProjectListFor(currentEmail);
      clearMatrixUI();
      updateProgressCounter();
      setStatus('Submitted! Select your next project or click "Submit ratings for project" when done.', 'success');

      if (hasCompletedAllProjects()) {
        setTimeout(showThankyouStage, 1000);
      }
    })
    .catch(function (err) {
      console.error('Submission error', err);
      if (err && err.alreadySubmitted) {
        completedProjects[submittingProject] = true;
        remoteCompletedProjects[submittingProject] = true;
        delete stagedRatings[submittingProject];
        saveProgress();
        populateProjectListFor(currentEmail);
        clearMatrixUI();
        updateProgressCounter();
        setStatus('This project was already submitted by another sponsor, so it is marked complete here.', 'info');
        if (hasCompletedAllProjects()) {
          setTimeout(showThankyouStage, 1000);
        }
        return;
      }
      setStatus('Submission failed. Check your connection and try again.', 'error');
    })
    .finally(function () {
      if (submitProjectBtn) submitProjectBtn.disabled = false;
    });
  }

  function clearMatrixUI() {
    if (matrixContainer) matrixContainer.innerHTML = '';
    var commentSec = document.querySelector('.section.section-comment');
    if (commentSec && commentSec.parentNode) commentSec.parentNode.removeChild(commentSec);
    var infoEl = $('matrix-info');
    if (infoEl && infoEl.parentNode) infoEl.parentNode.removeChild(infoEl);
    currentProject = '';
  }

  function hasCompletedAllProjects() {
    var entry = sponsorData[currentEmail] || {};
    var all   = Object.keys(entry.projects || {});
    if (!all.length) return false;
    for (var i = 0; i < all.length; i++) {
      if (!isProjectCompleted(all[i])) return false;
    }
    return true;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // REPORT GENERATION
  // ─────────────────────────────────────────────────────────────────────────────
  function generateReportHTML() {
    var h = '';
    h += '<!doctype html><html lang="en"><head>';
    h += '<meta charset="utf-8">';
    h += '<title>Sponsor Evaluation Report &ndash; ' + escapeHtml(currentName) + '</title>';
    h += '<style>';
    h += 'body{font-family:Arial,Helvetica,sans-serif;margin:40px;color:#1a1a2e;font-size:14px;}';
    h += 'h1{color:#8c1d40;border-bottom:3px solid #8c1d40;padding-bottom:8px;font-size:1.5rem;}';
    h += 'h2{color:#8c1d40;margin-top:36px;border-bottom:1px solid #ddd;padding-bottom:4px;font-size:1.15rem;}';
    h += 'h3{color:#333;margin-top:18px;font-size:1rem;}';
    h += '.meta{color:#666;font-size:0.88rem;margin:4px 0;}';
    h += 'table{width:100%;border-collapse:collapse;margin:12px 0;font-size:12px;}';
    h += 'th{background:#8c1d40;color:#fff;padding:7px 10px;text-align:left;font-weight:600;}';
    h += 'td{padding:7px 10px;border-bottom:1px solid #e8e8e8;vertical-align:top;}';
    h += 'tr:nth-child(even) td{background:#f9f9f9;}';
    h += '.comment-block{background:#f8f8f8;border-left:3px solid #8c1d40;padding:8px 12px;margin:6px 0;font-size:13px;white-space:pre-wrap;}';
    h += '.no-rating{color:#bbb;font-style:italic;}';
    h += '.badge-done{display:inline-block;background:#22c55e;color:#fff;border-radius:4px;padding:1px 7px;font-size:11px;font-weight:700;margin-left:8px;}';
    h += '@media print{button{display:none!important;}.no-print{display:none!important;}}';
    h += '</style></head><body>';

    // Print/download buttons (hidden in print)
    h += '<div class="no-print" style="margin-bottom:24px;">';
    h += '<button onclick="window.print()" style="margin-right:8px;padding:8px 18px;background:#8c1d40;color:#fff;border:0;border-radius:6px;cursor:pointer;font-size:14px;">Print</button>';
    h += '<button onclick="window.close()" style="padding:8px 18px;background:#eee;color:#333;border:0;border-radius:6px;cursor:pointer;font-size:14px;">Close</button>';
    h += '</div>';

    h += '<h1>CSE 485 Capstone Sponsor Evaluation Report</h1>';
    h += '<p class="meta"><strong>Sponsor:</strong> ' + escapeHtml(currentName) + ' &lt;' + escapeHtml(currentEmail) + '&gt;</p>';
    h += '<p class="meta"><strong>Survey Round:</strong> ' + escapeHtml(ROUND) + '</p>';
    h += '<p class="meta"><strong>Generated:</strong> ' + new Date().toLocaleString() + '</p>';

    var projectNames = Object.keys(submittedResponses);
    if (!projectNames.length) {
      h += '<p>No submissions recorded in this session. Submissions are only stored locally for report purposes.</p>';
    } else {
      projectNames.forEach(function (proj) {
        var payload = submittedResponses[proj];
        if (!payload) return;

        h += '<h2>' + escapeHtml(proj) + '<span class="badge-done">Submitted</span></h2>';
        h += '<p class="meta">Submitted: ' + new Date(payload.timestamp).toLocaleString() + '</p>';

        // Ratings table
        h += '<table><thead><tr><th>Student / Team</th>';
        RUBRIC.forEach(function (r) { h += '<th>' + escapeHtml(r.title) + '</th>'; });
        h += '</tr></thead><tbody>';
        payload.responses.forEach(function (resp) {
          h += '<tr><td><strong>' + escapeHtml(resp.student) + '</strong></td>';
          RUBRIC.forEach(function (r) {
            var val = resp.ratings[r.title];
            h += '<td>' + (val != null ? String(val) : '<span class="no-rating">&ndash;</span>') + '</td>';
          });
          h += '</tr>';
        });
        h += '</tbody></table>';

        // Comments
        payload.responses.forEach(function (resp) {
          var hasShared     = resp.commentShared     && resp.commentShared.trim();
          var hasInstructor = resp.commentInstructor && resp.commentInstructor.trim();
          if (!hasShared && !hasInstructor) return;

          h += '<h3>Comments &ndash; ' + escapeHtml(resp.student) + '</h3>';
          if (hasShared) {
            h += '<div class="comment-block"><strong>Shared with student:</strong><br>' +
                 escapeHtml(resp.commentShared) + '</div>';
          }
          if (hasInstructor) {
            h += '<div class="comment-block"><strong>Private (instructor only):</strong><br>' +
                 escapeHtml(resp.commentInstructor) + '</div>';
          }
        });
      });
    }

    h += '</body></html>';
    return h;
  }

  function downloadReport() {
    var html = generateReportHTML();
    var blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement('a');
    a.href     = url;
    a.download = 'sponsor-report-' +
                 currentName.replace(/\s+/g, '-').toLowerCase() +
                 '-' + ROUND + '.html';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function printReport() {
    var html = generateReportHTML();
    var win  = window.open('', '_blank', 'width=950,height=720,scrollbars=yes');
    if (!win) {
      alert('Please allow pop-ups to print the report.');
      return;
    }
    win.document.write(html);
    win.document.close();
    win.focus();
    // Small delay to ensure content is fully rendered before print dialog
    setTimeout(function () { win.print(); }, 600);
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // IDENTITY SUBMIT
  // ─────────────────────────────────────────────────────────────────────────────
  function onIdentitySubmit() {
    var name  = nameInput  ? nameInput.value.trim() : '';
    var email = emailInput ? emailInput.value.toLowerCase().trim() : '';

    if (!name) { setStatus('Please enter your name.', 'error'); return; }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      setStatus('Please enter a valid email address.', 'error');
      return;
    }

    currentName  = name;
    currentEmail = email;
    saveProgress();

    if (!Object.keys(sponsorData).length) {
      setStatus('Loading project data, please wait\u2026', 'info');
      if (identitySubmit) identitySubmit.disabled = true;
      tryFetchData(function () {
        if (identitySubmit) identitySubmit.disabled = false;
        if (!sponsorData[currentEmail]) {
          setStatus('No projects found for that email address. Please check and try again.', 'error');
          return;
        }
        showProjectsForCurrentEmail();
      });
    } else {
      if (!sponsorData[currentEmail]) {
        setStatus('No projects found for that email address. Please check and try again.', 'error');
        return;
      }
      showProjectsForCurrentEmail();
    }
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // EVENT WIRING
  // ─────────────────────────────────────────────────────────────────────────────

  // Allow Enter key to submit identity form
  if (nameInput)  nameInput.addEventListener('keydown',  function (e) { if (e.key === 'Enter') onIdentitySubmit(); });
  if (emailInput) emailInput.addEventListener('keydown', function (e) { if (e.key === 'Enter') onIdentitySubmit(); });

  if (identitySubmit)    identitySubmit.addEventListener('click', onIdentitySubmit);
  if (submitProjectBtn)  submitProjectBtn.addEventListener('click', submitCurrentProject);
  if (downloadReportBtn) downloadReportBtn.addEventListener('click', downloadReport);
  if (printReportBtn)    printReportBtn.addEventListener('click', printReport);

  if (backToIdentity) {
    backToIdentity.addEventListener('click', function () {
      clearMatrixUI();
      showIdentityStage();
    });
  }

  if (finishStartOverBtn) {
    finishStartOverBtn.addEventListener('click', function () {
      completedProjects  = {};
      stagedRatings      = {};
      submittedResponses = {};
      currentProject     = '';
      saveProgress();
      clearMatrixUI();
      showIdentityStage();
    });
  }

  // Warn before leaving if there are unsaved ratings
  window.addEventListener('beforeunload', function (e) {
    if (currentProject && Object.keys(stagedRatings[currentProject] || {}).length) {
      e.preventDefault();
      e.returnValue = 'You have unsaved ratings. Are you sure you want to leave?';
    }
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // STAGE DISPLAY HELPERS
  // ─────────────────────────────────────────────────────────────────────────────
  function showIdentityStage() {
    if (stageIdentity) stageIdentity.style.display = '';
    if (stageProjects) stageProjects.style.display = 'none';
    if (stageThankyou) stageThankyou.style.display = 'none';
    if (welcomeBlock)  welcomeBlock.style.display  = '';
    if (underTitle)    underTitle.style.display    = '';
    setStatus('');
  }

  function showProjectsStage() {
    if (stageIdentity) stageIdentity.style.display = 'none';
    if (stageProjects) stageProjects.style.display = '';
    if (stageThankyou) stageThankyou.style.display = 'none';
    if (welcomeBlock)  welcomeBlock.style.display  = 'none';
    if (underTitle)    underTitle.style.display    = 'none';
    setStatus('');
  }

  function showThankyouStage() {
    if (stageIdentity) stageIdentity.style.display = 'none';
    if (stageProjects) stageProjects.style.display = 'none';
    if (stageThankyou) stageThankyou.style.display = '';
    if (welcomeBlock)  welcomeBlock.style.display  = 'none';
    if (underTitle)    underTitle.style.display    = 'none';
    setStatus('');
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // DATA FETCHING
  // ─────────────────────────────────────────────────────────────────────────────
  function tryFetchData(callback) {
    fetch(DATA_LOADER_URL, { cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error('Data loader returned HTTP ' + r.status);
        return r.json();
      })
      .then(function (rows) {
        sponsorData = buildSponsorMap(rows || []);
        if (typeof callback === 'function') callback();
      })
      .catch(function (err) {
        console.error('Data fetch failed', err);
        setStatus('Could not load project data. Please refresh and try again.', 'error');
        if (typeof callback === 'function') callback();
      });
  }

  // ─────────────────────────────────────────────────────────────────────────────
  function fetchCompletionStatus(callback) {
    fetch(ENDPOINT_URL + '?status=1&round=' + encodeURIComponent(ROUND), { cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error('Completion status returned HTTP ' + r.status);
        return r.json();
      })
      .then(function (json) {
        var next = {};
        var source = (json && json.completedProjects) || {};
        if (Array.isArray(source)) {
          source.forEach(function (projectName) {
            if (projectName) next[projectName] = true;
          });
        } else {
          Object.keys(source).forEach(function (projectName) {
            if (source[projectName]) next[projectName] = true;
          });
        }
        remoteCompletedProjects = next;
        updateProgressCounter();
      })
      .catch(function (err) {
        console.info('Shared completion status unavailable', err);
      })
      .finally(function () {
        if (typeof callback === 'function') callback();
      });
  }

  function showProjectsForCurrentEmail() {
    fetchCompletionStatus(function () {
      showProjectsStage();
      populateProjectListFor(currentEmail);
      if (hasCompletedAllProjects()) {
        setStatus('All assigned projects have already been submitted. Completed projects are shown below.', 'info');
      }
    });
  }

  // BOOT SEQUENCE
  // ─────────────────────────────────────────────────────────────────────────────
  showIdentityStage();
  var hadProgress = loadProgress();  // pre-fills name/email inputs if progress exists

  // Load data in background; if user had saved progress, show a "welcome back" hint
  tryFetchData(function () {
    fetchCompletionStatus(function () {
      if (hadProgress && currentEmail && sponsorData[currentEmail]) {
        setStatus('Welcome back! Your previous progress has been restored. Click Continue to resume.', 'success');
      }
    });
  });

  // ─────────────────────────────────────────────────────────────────────────────
  // RADIO TOGGLE (click an already-selected radio to deselect it)
  // ─────────────────────────────────────────────────────────────────────────────
  (function () {
    function findRadio(e) {
      var path = (e.composedPath && e.composedPath()) || [];
      if (!path.length) {
        var node = e.target;
        while (node) { path.push(node); node = node.parentNode; }
      }
      for (var i = 0; i < path.length; i++) {
        var n = path[i];
        if (!n || !n.tagName) continue;
        var tag = n.tagName.toLowerCase();
        if (tag === 'input' && n.type === 'radio') return n;
        if (tag === 'label') {
          var q = n.querySelector && n.querySelector("input[type='radio']");
          if (q) return q;
          var fid = n.getAttribute && n.getAttribute('for');
          if (fid) { var byId = document.getElementById(fid); if (byId && byId.type === 'radio') return byId; }
        }
      }
      return null;
    }

    document.addEventListener('pointerdown', function (e) {
      try {
        var r = findRadio(e);
        if (r) r.dataset.waschecked = r.checked ? 'true' : 'false';
      } catch (_) {}
    }, false);

    document.addEventListener('touchstart', function (e) {
      try {
        var r = findRadio(e);
        if (r) r.dataset.waschecked = r.checked ? 'true' : 'false';
      } catch (_) {}
    }, { passive: true });

    document.addEventListener('keydown', function (e) {
      if (e.key !== ' ' && e.key !== 'Spacebar' && e.key !== 'Enter') return;
      var a = document.activeElement;
      if (a && a.tagName && a.tagName.toLowerCase() === 'input' && a.type === 'radio') {
        a.dataset.waschecked = a.checked ? 'true' : 'false';
      }
    }, false);

    document.addEventListener('click', function (e) {
      try {
        var r = findRadio(e);
        if (!r) return;
        if (r.dataset.waschecked === 'true') {
          Promise.resolve().then(function () {
            if (r.checked) {
              r.checked = false;
              r.dispatchEvent(new Event('change', { bubbles: true }));
            }
            r.removeAttribute('data-waschecked');
          });
        } else {
          r.dataset.waschecked = r.checked ? 'true' : 'false';
        }
      } catch (_) {}
    }, false);
  })();

  // ─────────────────────────────────────────────────────────────────────────────
  // DEBUG HELPERS (accessible from browser console)
  // ─────────────────────────────────────────────────────────────────────────────
  window.__sponsorDebug = {
    get sponsorData()        { return sponsorData; },
    get stagedRatings()      { return stagedRatings; },
    get completedProjects()  { return completedProjects; },
    get remoteCompletedProjects() { return remoteCompletedProjects; },
    get submittedResponses() { return submittedResponses; },
    get storageKey()         { return STORAGE_KEY; },
    reloadData:     tryFetchData,
    reloadCompletionStatus: fetchCompletionStatus,
    generateReport: generateReportHTML
  };

})();
