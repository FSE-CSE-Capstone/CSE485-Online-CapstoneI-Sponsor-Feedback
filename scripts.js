// CSE 485 Online Capstone I — Sponsor Feedback Survey
// Full fixed & featured version. Vanilla JS, single IIFE.
(function () {
  'use strict';

  // -------------------------------------------------------
  // Semester-aware localStorage key
  // Change window.SURVEY_ROUND in index.html each semester.
  // -------------------------------------------------------
  var ROUND = (window.SURVEY_ROUND || 'round1');
  var BASE_KEY = 'sponsor_progress_v1';
  var STORAGE_KEY = BASE_KEY + '_' + ROUND;

  // Remove stale keys from previous rounds (uses indexOf for broad compatibility)
  try {
    Object.keys(localStorage).forEach(function (k) {
      if (k.indexOf(BASE_KEY) === 0 && k !== STORAGE_KEY) {
        localStorage.removeItem(k);
      }
    });
  } catch (e) { console.warn('Round cleanup failed', e); }

  // -------------------------------------------------------
  // CONFIG — keep URLs exactly as-is
  // -------------------------------------------------------
  var ENDPOINT_URL    = 'https://cse485-online-worker.sbecerr7.workers.dev/';
  var DATA_LOADER_URL = 'https://cse485-online-data-loader.sbecerr7.workers.dev/';

  // -------------------------------------------------------
  // RUBRIC — keep exactly as-is
  // -------------------------------------------------------
  var RUBRIC = [
    { title: "Student has contributed an appropriate amount of development effort towards this project", description: "Development effort should be balanced between all team members; student should commit to a fair amount of development effort on each sprint." },
    { title: "Meetings", description: "Students are expected to be proactive. Contributions and participation in meetings help ensure the student is aware of project goals." },
    { title: "Understanding", description: "Students are expected to understand important details of the project and be able to explain it from different stakeholder perspectives." },
    { title: "Quality", description: "Students should complete assigned work to a high quality: correct, documented, and self-explanatory where appropriate." },
    { title: "Communication", description: "Students are expected to be in regular communication and maintain professionalism when interacting with the sponsor." }
  ];

  // -------------------------------------------------------
  // DOM refs
  // -------------------------------------------------------
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
  var finishStartOverBtn  = $('finishStartOver');
  var downloadReportBtn   = $('downloadReport');
  var printReportBtn      = $('printReport');
  var welcomeBlock     = $('welcome-block');
  var underTitle       = $('under-title');
  var progressCounter  = $('progress-counter');

  // -------------------------------------------------------
  // State
  // -------------------------------------------------------
  var sponsorData        = {};
  var sponsorProjects    = {};
  var currentEmail       = '';
  var currentName        = '';
  var currentProject     = '';
  var completedProjects  = {};
  var stagedRatings      = {};
  var submittedResponses = {};

  // -------------------------------------------------------
  // Helpers
  // -------------------------------------------------------

  // setStatus uses CSS type classes, not inline colors
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

  // Element builder utility
  function el(tag, props, children) {
    var n = document.createElement(tag);
    if (props) {
      Object.keys(props).forEach(function (k) {
        if (k === 'class')  n.className = props[k];
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

  // Update "X of Y projects completed" counter
  function updateProgressCounter() {
    if (!progressCounter) return;
    var entry = sponsorData[currentEmail] || {};
    var all = Object.keys(entry.projects || {});
    var done = 0;
    all.forEach(function (p) { if (completedProjects[p]) done++; });
    progressCounter.textContent = done + ' of ' + all.length + ' projects completed';
  }

  // -------------------------------------------------------
  // buildSponsorMap — kept exactly as-is
  // -------------------------------------------------------
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
        var rawVal = (rawRow[rawKey] || '').toString().replace(/\u00A0/g, ' ').trim();
        if (!project && /^(project|project name|project_title|group_name|projectname)$/.test(keyNorm)) project = rawVal;
        else if (!student && /^(student|student name|students|name|student_name)$/.test(keyNorm)) student = rawVal;
        else if (!sponsorCell && /^(sponsoremail|sponsor email|sponsor|email|login_id|sponsor_email)$/.test(keyNorm)) sponsorCell = rawVal;
      });

      if (!sponsorCell) {
        var fallback = [];
        Object.keys(rawRow || {}).forEach(function (k) {
          var rv = (rawRow[k] || '').toString();
          var found = rv.match(emailRegex);
          if (found) fallback = fallback.concat(found);
        });
        if (fallback.length) sponsorCell = fallback.join(', ');
      }

      project = (project || '').trim(); student = (student || '').trim();
      if (!sponsorCell || !project || !student) return;

      var tokens = sponsorCell.split(/[,;\/|]+/);
      var foundEmails = [];
      tokens.forEach(function (t) {
        var cleaned = cleanToken(t);
        if (!cleaned) return;
        var matches = cleaned.match(emailRegex) || t.match(emailRegex) || (t.replace(/\s+/g, '').match(emailRegex) || []);
        if (matches) matches.forEach(function (em) { foundEmails.push(em.toLowerCase().trim()); });
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
        if (map[email].projects[project].indexOf(student) === -1) map[email].projects[project].push(student);
      });
    });
    return map;
  }

  // -------------------------------------------------------
  // Persistence — includes submittedResponses
  // -------------------------------------------------------
  function saveProgress() {
    var payload = {
      name:               currentName,
      email:              currentEmail,
      completedProjects:  completedProjects,
      stagedRatings:      stagedRatings,
      submittedResponses: submittedResponses
    };
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(payload)); } catch (e) { console.warn('Could not save progress', e); }
  }

  // Returns true if saved progress was found and restored
  function loadProgress() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return false;
      var obj = JSON.parse(raw);
      if (obj && obj.email) {
        currentName         = obj.name  || '';
        currentEmail        = obj.email || '';
        completedProjects   = obj.completedProjects  || {};
        stagedRatings       = obj.stagedRatings      || {};
        submittedResponses  = obj.submittedResponses || {};
        if (nameInput)  nameInput.value  = currentName;
        if (emailInput) emailInput.value = currentEmail;
        return true;
      }
    } catch (e) { console.warn('Could not load progress', e); }
    return false;
  }

  // -------------------------------------------------------
  // Populate project list
  // -------------------------------------------------------
  function populateProjectListFor(email) {
    if (!projectListEl) return;
    projectListEl.innerHTML = '';
    sponsorProjects = {};

    var entry = sponsorData[email];
    if (!entry || !entry.projects) { setStatus('No projects found for that email.', 'error'); return; }

    var allProjects = Object.keys(entry.projects).slice();
    // Completed projects sort to the bottom
    allProjects.sort(function (a, b) {
      return (completedProjects[a] ? 1 : 0) - (completedProjects[b] ? 1 : 0);
    });

    allProjects.forEach(function (p) {
      var isDone = !!completedProjects[p];
      var li = el('li', {
        class: 'project-item' + (isDone ? ' completed' : ''),
        tabindex: isDone ? '-1' : '0',
        'data-project': p
      });

      if (isDone) {
        // Completed: show checkmark, no click/keydown listeners
        li.innerHTML = '&#10003; <strong>' + escapeHtml(p) + '</strong> <span class="meta">(Completed)</span>';
      } else {
        li.innerHTML = '<strong class="project-item-name">' + escapeHtml(p) + '</strong><span class="project-item-arrow">&#8250;</span>';
        li.addEventListener('click', function () {
          Array.from(projectListEl.querySelectorAll('.project-item.active')).forEach(function (a) { a.classList.remove('active'); });
          li.classList.add('active');
          currentProject = p;
          loadProjectIntoMatrix(p, entry.projects[p]);
          setStatus('');
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

  // Remove empty placeholder cards (defensive)
  function removeEmptyPlaceholderCards() {
    if (!projectListEl) return;
    var container = projectListEl.parentNode;
    if (!container) return;
    Array.from(container.querySelectorAll('.card')).forEach(function (c) {
      var hasControls = c.querySelector('input, textarea, select, button, table, label');
      var text = (c.textContent || '').replace(/\s+/g, '');
      if (!hasControls && text.length === 0 && !c.classList.contains('matrix-card') && !c.classList.contains('persistent-placeholder')) {
        c.parentNode && c.parentNode.removeChild(c);
      }
    });
  }

  // -------------------------------------------------------
  // Validate ratings before submission
  // -------------------------------------------------------
  function validateRatings(students) {
    var issues = [];
    for (var c = 0; c < RUBRIC.length; c++) {
      // Pass if team row covers this criterion
      if (document.querySelector('input[name="rating-' + c + '-team"]:checked')) continue;
      // Otherwise at least one student must be rated
      var anyStudentRated = false;
      for (var s = 0; s < students.length; s++) {
        if (document.querySelector('input[name="rating-' + c + '-' + s + '"]:checked')) {
          anyStudentRated = true;
          break;
        }
      }
      if (!anyStudentRated) {
        issues.push('Please rate criterion ' + (c + 1) + ': "' + RUBRIC[c].title + '"');
      }
    }
    return issues;
  }

  // -------------------------------------------------------
  // Build matrix for a project
  // -------------------------------------------------------
  function loadProjectIntoMatrix(projectName, students) {
    if (!projectName) return;
    currentProject = projectName;

    // Cleanup prior UI fragments
    var existingInfo = $('matrix-info');
    if (existingInfo && existingInfo.parentNode) existingInfo.parentNode.removeChild(existingInfo);
    Array.from(document.querySelectorAll('.current-project-header')).forEach(function (h) {
      if (h.parentNode) h.parentNode.removeChild(h);
    });
    var oldComment = document.querySelector('.section.section-comment');
    if (oldComment && oldComment.parentNode) oldComment.parentNode.removeChild(oldComment);

    // Header / info block
    var info = el('div', { id: 'matrix-info', class: 'matrix-info-block' });
    info.appendChild(el('div', { class: 'current-project-header', text: projectName }));
    info.appendChild(el('div', { class: 'matrix-info-desc', text: 'Please evaluate the students using the rubric below (scale 1\u20137).' }));
    if (matrixContainer && matrixContainer.parentNode) {
      matrixContainer.parentNode.insertBefore(info, matrixContainer);
    } else if (matrixContainer) {
      document.body.insertBefore(info, matrixContainer);
    }

    if (!students || !students.length) {
      if (matrixContainer) matrixContainer.textContent = 'No students found for this project.';
      return;
    }

    if (!stagedRatings[currentProject]) stagedRatings[currentProject] = {};

    var temp = document.createElement('div');

    RUBRIC.forEach(function (crit, cIdx) {
      var card = el('div', { class: 'card matrix-card', style: { marginBottom: '20px', padding: '18px' } });
      var critWrap = el('div', { class: 'matrix-criterion' });
      critWrap.appendChild(el('h4', { class: 'matrix-criterion-title', text: (cIdx + 1) + '. ' + crit.title }));
      critWrap.appendChild(el('div', { class: 'matrix-criterion-desc', text: crit.description }));

      var scrollWrap = el('div', { class: 'table-scroll-wrap' });
      var table = el('table', { class: 'matrix-table' });

      var colgroup = el('colgroup');
      colgroup.appendChild(el('col', { style: { width: '46%' } }));
      colgroup.appendChild(el('col', { style: { width: '12%' } }));
      for (var ci = 0; ci < 7; ci++) colgroup.appendChild(el('col', { style: { width: '4%' } }));
      colgroup.appendChild(el('col', { style: { width: '12%' } }));
      table.appendChild(colgroup);

      var thead = el('thead');
      var trHead = el('tr');
      trHead.appendChild(el('th', { text: 'Student', style: { textAlign: 'left', padding: '8px' } }));
      trHead.appendChild(el('th', { class: 'header-descriptor', html: '<div class="hd-line">Far Below Expectations</div><div class="hd-sub">(Fail)</div>' }));
      for (var k = 1; k <= 7; k++) trHead.appendChild(el('th', { class: 'col-score-num', text: String(k) }));
      trHead.appendChild(el('th', { class: 'header-descriptor header-descriptor-right', html: '<div class="hd-line">Exceeds Expectations</div><div class="hd-sub">(A+)</div>' }));
      thead.appendChild(trHead);
      table.appendChild(thead);

      var tbody = el('tbody');

      students.forEach(function (studentName, sIdx) {
        var tr = el('tr', { class: sIdx % 2 === 0 ? 'row-even' : 'row-odd' });
        tr.appendChild(el('td', { class: 'col-student', text: studentName }));
        tr.appendChild(el('td', { class: 'col-descriptor' }));

        var stagedForStudent = (stagedRatings[currentProject] && stagedRatings[currentProject][sIdx]) || {};
        for (var score = 1; score <= 7; score++) {
          var td = el('td', { class: 'col-radio' });
          var inputId = 'rating-' + cIdx + '-' + sIdx + '-' + score;
          var input = el('input', { type: 'radio', name: 'rating-' + cIdx + '-' + sIdx, value: String(score), id: inputId });
          if (stagedForStudent[cIdx] && String(stagedForStudent[cIdx]) === String(score)) input.checked = true;
          var label = el('label', { class: 'radio-label', 'for': inputId });
          label.appendChild(input);
          td.appendChild(label);
          tr.appendChild(td);
        }

        tr.appendChild(el('td', { class: 'col-descriptor' }));
        tbody.appendChild(tr);
      });

      // Team overall row
      var trTeam = el('tr', { class: 'row-team' });
      trTeam.appendChild(el('td', { class: 'col-student', text: 'Team Overall' }));
      trTeam.appendChild(el('td', { class: 'col-descriptor' }));
      var stagedTeam = (stagedRatings[currentProject] && stagedRatings[currentProject].team) || {};
      for (var sScore = 1; sScore <= 7; sScore++) {
        var tdT = el('td', { class: 'col-radio' });
        var teamInputId = 'rating-' + cIdx + '-team-' + sScore;
        var inputT = el('input', { type: 'radio', name: 'rating-' + cIdx + '-team', value: String(sScore), id: teamInputId });
        if (stagedTeam[cIdx] && String(stagedTeam[cIdx]) === String(sScore)) inputT.checked = true;
        var lblT = el('label', { class: 'radio-label', 'for': teamInputId });
        lblT.appendChild(inputT);
        tdT.appendChild(lblT);
        trTeam.appendChild(tdT);
      }
      trTeam.appendChild(el('td', { class: 'col-descriptor' }));
      tbody.appendChild(trTeam);

      table.appendChild(tbody);
      scrollWrap.appendChild(table);
      critWrap.appendChild(scrollWrap);
      card.appendChild(critWrap);
      temp.appendChild(card);
    });

    if (matrixContainer) {
      matrixContainer.innerHTML = '';
      while (temp.firstChild) matrixContainer.appendChild(temp.firstChild);
    }

    // renderCommentSection first, then attach matrix listeners
    renderCommentSection(projectName, students);
    attachMatrixListeners();

    if (typeof window.__attachRadioToggle === 'function') {
      Array.prototype.forEach.call(matrixContainer.querySelectorAll("input[type='radio']"), function (r) {
        window.__attachRadioToggle(r);
      });
    }

    // Auto-scroll to matrix info block after DOM renders
    setTimeout(function () {
      var infoEl = $('matrix-info');
      if (infoEl) infoEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 120);
  }

  // -------------------------------------------------------
  // Render comment section
  // Textarea input listeners attached here (not in attachMatrixListeners)
  // -------------------------------------------------------
  function renderCommentSection(projectName, students) {
    var oldComment = document.querySelector('.section.section-comment');
    if (oldComment && oldComment.parentNode) oldComment.parentNode.removeChild(oldComment);

    var commentSec = el('div', { class: 'section section-comment' });
    commentSec.appendChild(el('h3', { class: 'comment-section-title', text: 'Add your additional comments' }));

    var staged = (stagedRatings[projectName] && stagedRatings[projectName]._studentComments) || {};

    students.forEach(function (studentName, sIdx) {
      var wrapper = el('div', { class: 'student-comment-panel' });
      var headerRow = el('div', { class: 'comment-panel-header' });
      headerRow.appendChild(el('div', { class: 'comment-panel-name', text: studentName }));

      var toggleBtn = el('button', { type: 'button', class: 'btn btn-mini comment-toggle', text: '\u25be Add comment' });
      headerRow.appendChild(toggleBtn);
      wrapper.appendChild(headerRow);

      var content = el('div', { class: 'student-comment-content', style: { display: 'none' } });
      content.appendChild(el('div', { class: 'comment-label', text: 'Comments to be SHARED WITH THE STUDENT' }));
      var taPublic = el('textarea', { id: 'comment-public-' + sIdx, placeholder: 'Comments to share with student', style: { width: '100%', minHeight: '60px', padding: '8px', boxSizing: 'border-box', marginBottom: '8px' } });
      content.appendChild(taPublic);
      content.appendChild(el('div', { class: 'comment-label', text: 'Comments to be SHARED ONLY WITH THE INSTRUCTOR' }));
      var taPrivate = el('textarea', { id: 'comment-private-' + sIdx, placeholder: 'Private comments for instructor', style: { width: '100%', minHeight: '60px', padding: '8px', boxSizing: 'border-box' } });
      content.appendChild(taPrivate);

      // Auto-save listeners on textareas — attached here so they are never missed
      taPublic.addEventListener('input', saveDraftHandler);
      taPrivate.addEventListener('input', saveDraftHandler);

      toggleBtn.addEventListener('click', function () {
        if (content.style.display === 'none') {
          content.style.display = 'block'; toggleBtn.textContent = '\u25b4 Hide comment';
        } else {
          content.style.display = 'none'; toggleBtn.textContent = '\u25be Add comment';
        }
      });

      var st = staged && staged[studentName];
      if (st) {
        if (st.public)  taPublic.value  = st.public;
        if (st.private) taPrivate.value = st.private;
        if ((st.public && st.public.length) || (st.private && st.private.length)) {
          content.style.display = 'block'; toggleBtn.textContent = '\u25b4 Hide comment';
        }
      }

      wrapper.appendChild(content);
      commentSec.appendChild(wrapper);
    });

    // Group / team overall panel
    var groupWrap = el('div', { class: 'student-comment-panel' });
    var groupHeader = el('div', { class: 'comment-panel-header' });
    groupHeader.appendChild(el('div', { class: 'comment-panel-name', text: 'Comments for team overall' }));
    var groupToggle = el('button', { type: 'button', class: 'btn btn-mini comment-toggle', text: '\u25be Add comment' });
    groupHeader.appendChild(groupToggle);
    groupWrap.appendChild(groupHeader);

    var groupContent = el('div', { style: { display: 'none' } });
    groupContent.appendChild(el('div', { class: 'comment-label', text: 'Comments for team overall (shared with student by default)' }));
    var taGroup = el('textarea', { id: 'comment-group-public', placeholder: 'Comments for team overall', style: { width: '100%', minHeight: '80px', padding: '8px', boxSizing: 'border-box' } });
    groupContent.appendChild(taGroup);
    groupContent.appendChild(el('div', { class: 'comment-label', style: { margin: '8px 0 4px 0' }, text: 'Private comments about the team (instructor only)' }));
    var taGroupPrivate = el('textarea', { id: 'comment-group-private', placeholder: 'Private comments for instructor about the team', style: { width: '100%', minHeight: '60px', padding: '8px', boxSizing: 'border-box' } });
    groupContent.appendChild(taGroupPrivate);

    // Auto-save listeners
    taGroup.addEventListener('input', saveDraftHandler);
    taGroupPrivate.addEventListener('input', saveDraftHandler);

    groupToggle.addEventListener('click', function () {
      if (groupContent.style.display === 'none') {
        groupContent.style.display = 'block'; groupToggle.textContent = '\u25b4 Hide comment';
      } else {
        groupContent.style.display = 'none'; groupToggle.textContent = '\u25be Add comment';
      }
    });

    var stagedGroup = (stagedRatings[currentProject] && stagedRatings[currentProject]._groupComments) || {};
    if (stagedGroup) {
      if (stagedGroup.public)  taGroup.value        = stagedGroup.public;
      if (stagedGroup.private) taGroupPrivate.value = stagedGroup.private;
      if ((stagedGroup.public && stagedGroup.public.length) || (stagedGroup.private && stagedGroup.private.length)) {
        groupContent.style.display = 'block'; groupToggle.textContent = '\u25b4 Hide comment';
      }
    }

    groupWrap.appendChild(groupContent);
    commentSec.appendChild(groupWrap);

    if (matrixContainer && matrixContainer.parentNode) {
      if (matrixContainer.nextSibling) matrixContainer.parentNode.insertBefore(commentSec, matrixContainer.nextSibling);
      else matrixContainer.parentNode.appendChild(commentSec);
    } else {
      document.body.appendChild(commentSec);
    }
  }

  // -------------------------------------------------------
  // Attach matrix change/input listeners (radio buttons only)
  // Textarea listeners are wired inside renderCommentSection
  // -------------------------------------------------------
  function attachMatrixListeners() {
    if (!matrixContainer) return;
    matrixContainer.removeEventListener('change', saveDraftHandler);
    matrixContainer.removeEventListener('input',  saveDraftHandler);
    matrixContainer.addEventListener('change', saveDraftHandler);
    matrixContainer.addEventListener('input',  saveDraftHandler);
  }

  // -------------------------------------------------------
  // Save draft — collect ratings and comments into stagedRatings
  // -------------------------------------------------------
  function saveDraftHandler() {
    if (!currentProject) return;
    if (!stagedRatings[currentProject]) stagedRatings[currentProject] = {};
    var students = sponsorProjects[currentProject] || [];

    for (var s = 0; s < students.length; s++) {
      stagedRatings[currentProject][s] = stagedRatings[currentProject][s] || {};
      for (var c = 0; c < RUBRIC.length; c++) {
        var sel = document.querySelector('input[name="rating-' + c + '-' + s + '"]:checked');
        if (sel) stagedRatings[currentProject][s][c] = parseInt(sel.value, 10);
        else if (stagedRatings[currentProject][s][c] !== undefined) { /* keep existing */ }
        else stagedRatings[currentProject][s][c] = null;
      }
    }

    stagedRatings[currentProject].team = stagedRatings[currentProject].team || {};
    for (var ct = 0; ct < RUBRIC.length; ct++) {
      var selT = document.querySelector('input[name="rating-' + ct + '-team"]:checked');
      if (selT) stagedRatings[currentProject].team[ct] = parseInt(selT.value, 10);
      else if (stagedRatings[currentProject].team[ct] !== undefined) { /* keep */ }
      else stagedRatings[currentProject].team[ct] = null;
    }

    stagedRatings[currentProject]._studentComments = stagedRatings[currentProject]._studentComments || {};
    for (var i = 0; i < students.length; i++) {
      var sName  = students[i];
      var pubEl  = document.getElementById('comment-public-'  + i);
      var privEl = document.getElementById('comment-private-' + i);
      stagedRatings[currentProject]._studentComments[sName] = stagedRatings[currentProject]._studentComments[sName] || { public: '', private: '' };
      if (pubEl)  stagedRatings[currentProject]._studentComments[sName].public  = pubEl.value  || '';
      if (privEl) stagedRatings[currentProject]._studentComments[sName].private = privEl.value || '';
    }

    stagedRatings[currentProject]._groupComments = stagedRatings[currentProject]._groupComments || { public: '', private: '' };
    var gpPub  = document.getElementById('comment-group-public');
    var gpPriv = document.getElementById('comment-group-private');
    if (gpPub)  stagedRatings[currentProject]._groupComments.public  = gpPub.value  || '';
    if (gpPriv) stagedRatings[currentProject]._groupComments.private = gpPriv.value || '';

    saveProgress();
  }

  // -------------------------------------------------------
  // Submit current project
  // -------------------------------------------------------
  function submitCurrentProject() {
    if (!currentProject) { setStatus('No project is loaded.', 'error'); return; }
    var students = sponsorProjects[currentProject] || [];
    if (!students.length) { setStatus('No students to submit.', 'error'); return; }

    // Validate — abort if any criterion is uncovered
    var issues = validateRatings(students);
    if (issues.length) { setStatus(issues[0], 'error'); return; }

    var responses = [];
    for (var s = 0; s < students.length; s++) {
      var ratingsObj = {};
      for (var c = 0; c < RUBRIC.length; c++) {
        var sel = document.querySelector('input[name="rating-' + c + '-' + s + '"]:checked');
        ratingsObj[RUBRIC[c].title || ('C' + c)] = sel ? parseInt(sel.value, 10) : null;
      }
      var commentShared     = (document.getElementById('comment-public-'  + s) || {}).value || '';
      var commentInstructor = (document.getElementById('comment-private-' + s) || {}).value || '';
      responses.push({ student: students[s], ratings: ratingsObj, commentShared: commentShared, commentInstructor: commentInstructor, isTeam: false });
    }

    var teamRatingsChosen = false;
    var teamRatingsObj = {};
    for (var tc = 0; tc < RUBRIC.length; tc++) {
      var teamSel = document.querySelector('input[name="rating-' + tc + '-team"]:checked');
      teamRatingsObj[RUBRIC[tc].title || ('C' + tc)] = teamSel ? parseInt(teamSel.value, 10) : null;
      if (teamSel) teamRatingsChosen = true;
    }
    var groupCommentShared     = (document.getElementById('comment-group-public')  || {}).value || '';
    var groupCommentInstructor = (document.getElementById('comment-group-private') || {}).value || '';
    if (teamRatingsChosen || groupCommentShared || groupCommentInstructor) {
      responses.push({ student: 'Evaluating group as a whole', ratings: teamRatingsObj, commentShared: groupCommentShared, commentInstructor: groupCommentInstructor, isTeam: true });
    }

    if (!responses.length) { setStatus('Nothing to submit.', 'error'); return; }

    var payload = {
      sponsorName:  currentName  || (nameInput  ? nameInput.value.trim()  : ''),
      sponsorEmail: currentEmail || (emailInput ? emailInput.value.trim() : ''),
      project:      currentProject,
      round:        ROUND,
      rubric:       RUBRIC.map(function (r) { return r.title; }),
      responses:    responses,
      timestamp:    new Date().toISOString()
    };

    // Store BEFORE fetch so report is available even if network is slow
    submittedResponses[currentProject] = payload;
    saveProgress();

    setStatus('Submitting\u2026', 'info');
    if (submitProjectBtn) submitProjectBtn.disabled = true;

    fetch(ENDPOINT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).then(function (resp) {
      if (!resp.ok) {
        return resp.text().then(function (txt) { throw new Error('Server error ' + resp.status + ': ' + txt); });
      }
      return resp.json().catch(function () { return {}; });
    }).then(function () {
      setStatus('Submission saved. Thank you!', 'success');

      completedProjects[currentProject] = true;
      if (stagedRatings[currentProject]) delete stagedRatings[currentProject];
      saveProgress();
      updateProgressCounter();

      // Update list item — strip listeners by replacing node
      if (projectListEl) {
        var escapedProject = (window.CSS && CSS.escape) ? CSS.escape(currentProject) : currentProject.replace(/["\\]/g, '\\$&');
        var li = projectListEl.querySelector('li[data-project="' + escapedProject + '"]');
        if (li) {
          li.classList.add('completed');
          li.classList.remove('active');
          li.setAttribute('tabindex', '-1');
          li.innerHTML = '&#10003; <strong>' + escapeHtml(currentProject) + '</strong> <span class="meta">(Completed)</span>';
          // Replace node to remove all event listeners
          var newLi = li.cloneNode(true);
          li.parentNode.replaceChild(newLi, li);
        }
      }

      // Cleanup matrix UI
      if (matrixContainer) matrixContainer.innerHTML = '';
      var commentSection = document.querySelector('.section.section-comment');
      if (commentSection && commentSection.parentNode) commentSection.parentNode.removeChild(commentSection);
      var matrixInfoBlock = $('matrix-info');
      if (matrixInfoBlock && matrixInfoBlock.parentNode) matrixInfoBlock.parentNode.removeChild(matrixInfoBlock);

      currentProject = '';
      if (hasCompletedAllProjects()) showThankyouStage();

    }).catch(function (err) {
      console.error('Submission failed', err);
      setStatus('Submission failed. Please try again or check your connection.', 'error');
    }).finally(function () {
      if (submitProjectBtn) submitProjectBtn.disabled = false;
    });
  }

  function hasCompletedAllProjects() {
    var entry = sponsorData[currentEmail] || {};
    var all = Object.keys(entry.projects || {});
    if (!all.length) return false;
    for (var i = 0; i < all.length; i++) { if (!completedProjects[all[i]]) return false; }
    return true;
  }

  // -------------------------------------------------------
  // Report generation
  // -------------------------------------------------------
  function generateReportHTML() {
    var h = [];
    h.push('<!doctype html><html lang="en"><head><meta charset="utf-8">');
    h.push('<title>Sponsor Report \u2014 ' + escapeHtml(currentName) + '</title>');
    h.push('<style>');
    h.push('body{font-family:system-ui,Arial,sans-serif;padding:24px;color:#07122a;max-width:1100px;margin:0 auto;}');
    h.push('h1,h2,h3{color:#8c1d40;}');
    h.push('table{border-collapse:collapse;width:100%;margin-bottom:16px;}');
    h.push('th,td{border:1px solid #ddd;padding:8px 10px;text-align:left;font-size:0.93rem;}');
    h.push('th{background:#f3f4f6;font-weight:700;}');
    h.push('.project-section{margin-bottom:32px;page-break-inside:avoid;}');
    h.push('.comment-block{background:#fafafa;border:1px solid #e6e6e6;border-radius:6px;padding:10px;margin:4px 0;font-size:0.9rem;}');
    h.push('.print-hide{margin-bottom:20px;}');
    h.push('@media print{.print-hide{display:none!important;}}');
    h.push('</style></head><body>');
    h.push('<div class="print-hide">');
    h.push('<button onclick="window.print()" style="margin-right:8px;">Print</button>');
    h.push('<button onclick="window.close()">Close</button>');
    h.push('</div>');
    h.push('<h1>Sponsor Evaluation Report</h1>');
    h.push('<p><strong>Sponsor:</strong> ' + escapeHtml(currentName) + ' &lt;' + escapeHtml(currentEmail) + '&gt;</p>');
    h.push('<p><strong>Survey Round:</strong> ' + escapeHtml(ROUND) + '</p>');
    h.push('<p><strong>Generated:</strong> ' + new Date().toLocaleString() + '</p>');
    h.push('<hr>');

    var projects = Object.keys(submittedResponses);
    if (!projects.length) {
      h.push('<p><em>No submissions recorded in this session.</em></p>');
    } else {
      projects.forEach(function (proj) {
        var p = submittedResponses[proj];
        h.push('<div class="project-section">');
        h.push('<h2>' + escapeHtml(proj) + '</h2>');
        h.push('<p><em>Submitted: ' + escapeHtml(p.timestamp || '') + '</em></p>');
        var rubricTitles = p.rubric || [];
        h.push('<table><thead><tr><th>Student</th>');
        rubricTitles.forEach(function (t) { h.push('<th>' + escapeHtml(t) + '</th>'); });
        h.push('</tr></thead><tbody>');
        (p.responses || []).forEach(function (r) {
          h.push('<tr><td><strong>' + escapeHtml(r.student) + '</strong></td>');
          rubricTitles.forEach(function (t) {
            var val = r.ratings ? r.ratings[t] : null;
            h.push('<td>' + (val != null ? escapeHtml(String(val)) : '&mdash;') + '</td>');
          });
          h.push('</tr>');
          if (r.commentShared) {
            h.push('<tr><td colspan="' + (rubricTitles.length + 1) + '">');
            h.push('<div class="comment-block"><strong>Shared comment:</strong> ' + escapeHtml(r.commentShared) + '</div></td></tr>');
          }
          if (r.commentInstructor) {
            h.push('<tr><td colspan="' + (rubricTitles.length + 1) + '">');
            h.push('<div class="comment-block"><strong>Private comment (instructor only):</strong> ' + escapeHtml(r.commentInstructor) + '</div></td></tr>');
          }
        });
        h.push('</tbody></table></div>');
      });
    }
    h.push('</body></html>');
    return h.join('\n');
  }

  function downloadReport() {
    var html = generateReportHTML();
    var blob = new Blob([html], { type: 'text/html' });
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement('a');
    a.href   = url;
    a.download = 'sponsor-report-' + (currentName || 'report').replace(/\s+/g, '-') + '-' + ROUND + '.html';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
  }

  function printReport() {
    var html = generateReportHTML();
    var win  = window.open('', '_blank', 'width=960,height=720');
    if (!win) { alert('Popup blocked. Please allow popups for this site to print the report.'); return; }
    win.document.open();
    win.document.write(html);
    win.document.close();
    setTimeout(function () { win.print(); }, 600);
  }

  // -------------------------------------------------------
  // Identity submit
  // -------------------------------------------------------
  function onIdentitySubmit() {
    var name  = nameInput  ? nameInput.value.trim()  : '';
    var email = emailInput ? (emailInput.value || '').toLowerCase().trim() : '';
    if (!name)  { setStatus('Please enter your name.',          'error'); return; }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { setStatus('Please enter a valid email.', 'error'); return; }

    currentName = name; currentEmail = email; saveProgress();

    if (!sponsorData || Object.keys(sponsorData).length === 0) {
      setStatus('Loading project data, please wait\u2026', 'info');
      tryFetchData(function () {
        if (!sponsorData || !sponsorData[currentEmail]) { setStatus('No projects found for that email.', 'error'); return; }
        showProjectsStage();
        populateProjectListFor(currentEmail);
      });
    } else {
      if (!sponsorData[currentEmail]) { setStatus('No projects found for that email.', 'error'); return; }
      showProjectsStage();
      populateProjectListFor(currentEmail);
    }
  }

  // -------------------------------------------------------
  // Event wiring — each button/input wired exactly once
  // -------------------------------------------------------
  if (identitySubmit)  identitySubmit.addEventListener('click', onIdentitySubmit);
  if (backToIdentity)  backToIdentity.addEventListener('click', showIdentityStage);
  if (submitProjectBtn) submitProjectBtn.addEventListener('click', submitCurrentProject);
  if (downloadReportBtn) downloadReportBtn.addEventListener('click', downloadReport);
  if (printReportBtn)    printReportBtn.addEventListener('click',    printReport);

  // Enter key on identity form
  if (nameInput)  nameInput.addEventListener('keydown',  function (e) { if (e.key === 'Enter') onIdentitySubmit(); });
  if (emailInput) emailInput.addEventListener('keydown', function (e) { if (e.key === 'Enter') onIdentitySubmit(); });

  if (finishStartOverBtn) finishStartOverBtn.addEventListener('click', function () {
    completedProjects  = {};
    stagedRatings      = {};
    submittedResponses = {};
    currentProject     = '';
    saveProgress();
    if (matrixContainer) matrixContainer.innerHTML = '';
    var cs = document.querySelector('.section.section-comment');
    if (cs && cs.parentNode) cs.parentNode.removeChild(cs);
    var mi = $('matrix-info');
    if (mi && mi.parentNode) mi.parentNode.removeChild(mi);
    showIdentityStage();
  });

  // Unsaved-changes warning
  window.addEventListener('beforeunload', function (e) {
    if (currentProject && stagedRatings[currentProject] && Object.keys(stagedRatings[currentProject]).length > 0) {
      e.preventDefault();
      e.returnValue = 'You have unsaved ratings for the current project. Are you sure you want to leave?';
    }
  });

  // -------------------------------------------------------
  // Stage display helpers — every helper sets ALL toggled elements
  // -------------------------------------------------------
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
  }

  function showThankyouStage() {
    if (stageIdentity) stageIdentity.style.display = 'none';
    if (stageProjects) stageProjects.style.display = 'none';
    if (stageThankyou) stageThankyou.style.display = '';
    if (welcomeBlock)  welcomeBlock.style.display  = 'none';
    if (underTitle)    underTitle.style.display    = 'none';
  }

  // -------------------------------------------------------
  // Fetch sponsor data
  // Does NOT call loadProgress() or auto-advance stage
  // -------------------------------------------------------
  function tryFetchData(callback) {
    console.info('tryFetchData: requesting', DATA_LOADER_URL);
    fetch(DATA_LOADER_URL, { cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error('Data loader returned ' + r.status);
        return r.json();
      })
      .then(function (rows) {
        sponsorData = buildSponsorMap(rows || []);
        if (typeof callback === 'function') callback();
      })
      .catch(function (err) {
        console.error('Data fetch failed', err);
        setStatus('Project data not found. Please try again later.', 'error');
        if (typeof callback === 'function') callback();
      });
  }

  // -------------------------------------------------------
  // Boot sequence
  // -------------------------------------------------------
  showIdentityStage();
  var hadProgress = loadProgress();  // restore state, pre-fill inputs
  tryFetchData(function () {
    // Welcome back message — do NOT auto-advance; user must click Continue
    if (hadProgress && currentEmail && sponsorData[currentEmail]) {
      setStatus('Welcome back! Your previous progress has been restored. Click Continue to resume.', 'success');
    }
  });

  // UI cleanup on DOM ready (defensive)
  document.addEventListener('DOMContentLoaded', function () {
    var autoFooter = document.querySelector('.site-footer-fixed');
    if (autoFooter) autoFooter.parentNode.removeChild(autoFooter);
  });

  // -------------------------------------------------------
  // Debug helpers
  // -------------------------------------------------------
  window.__sponsorDebug = {
    get sponsorData()        { return sponsorData; },
    get stagedRatings()      { return stagedRatings; },
    get completedProjects()  { return completedProjects; },
    get submittedResponses() { return submittedResponses; },
    get storageKey()         { return STORAGE_KEY; },
    reloadData:     tryFetchData,
    generateReport: generateReportHTML
  };
  window.__submitCurrentProject = submitCurrentProject;

  // -------------------------------------------------------
  // Single robust radio-toggle implementation (kept as-is)
  // -------------------------------------------------------
  (function () {
    function findRadioFromEvent(e) {
      var path = (e.composedPath && e.composedPath()) || e.path;
      if (!path) {
        path = [];
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
          var forId = n.getAttribute && n.getAttribute('for');
          if (forId) {
            var byId = document.getElementById(forId);
            if (byId && byId.type === 'radio') return byId;
          }
        }
      }
      return null;
    }

    document.addEventListener('pointerdown', function (e) {
      try {
        var radio = findRadioFromEvent(e);
        if (!radio) return;
        radio.dataset.waschecked = radio.checked ? 'true' : 'false';
      } catch (err) {}
    }, false);

    document.addEventListener('touchstart', function (e) {
      try {
        var radio = findRadioFromEvent(e);
        if (!radio) return;
        radio.dataset.waschecked = radio.checked ? 'true' : 'false';
      } catch (err) {}
    }, { passive: true });

    document.addEventListener('keydown', function (e) {
      if (e.key !== ' ' && e.key !== 'Spacebar' && e.key !== 'Enter') return;
      var active = document.activeElement;
      if (!active) return;
      if (active.tagName && active.tagName.toLowerCase() === 'input' && active.type === 'radio') {
        active.dataset.waschecked = active.checked ? 'true' : 'false';
      }
    }, false);

    document.addEventListener('click', function (e) {
      try {
        var radio = findRadioFromEvent(e);
        if (!radio) return;
        if (radio.dataset.waschecked === 'true') {
          Promise.resolve().then(function () {
            if (radio.checked) {
              radio.checked = false;
              radio.dispatchEvent(new Event('change', { bubbles: true }));
            }
            radio.removeAttribute('data-waschecked');
          });
          return;
        }
        radio.dataset.waschecked = radio.checked ? 'true' : 'false';
      } catch (err) {
        console.error('radio-toggle error', err);
      }
    }, false);
  })();

  window.__attachRadioToggle = function (radio) {
    try {
      if (!radio || radio.dataset.toggleAttached === '1') return;
      radio.dataset.toggleAttached = '1';
    } catch (e) {}
  };

})();

