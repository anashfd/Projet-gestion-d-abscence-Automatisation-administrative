
/* ---------- items d list d excel 3la had chkel 
{
  MatriculeEtudiant: "matricule",
  Nom: "nom",
  Prenom: "prenom",
  CIN: "cin",
  LibelleLong: "e.g AG_INFO_TS_1A-Infographie (1A)-2025",
  CodeDiplome: "e.g INFO102",
  Nom_Arabe: "كنية",
  Prenom_arabe: "اسم",
 ... (DateNaissance, Adresse, etc.)
}
*/
const allRows = [];


// ------------- excel file reading 

document.getElementById("file").onchange = () => {

  const reader = new FileReader();

  reader.onload = () => {

    const workbook = XLSX.read(reader.result, { type: "array" });
    const sheet = workbook.Sheets["Export"];
    const sheetRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    // ka nredo allRows khawya (length = 0) 3la 9bel la bghina nbedlo file
    allRows.length = 0;
    // kankhdmo ghir b les lignes li fihom CodeDiplome w Nom 
    for (const row of sheetRows) {
      if (row.CodeDiplome !== "" && row.Nom !== "") {
        allRows.push(row);
      }
    }

    // lista dyal les groupes sans doublants and group libre
    const groups = [];
    for (const row of allRows) {
      const name = row.CodeDiplome;
      if (!groups.includes(name) && name !== "Groupe libre") {
        groups.push(name);
      }
    }
    groups.sort();

    let optionsHtml = "";
    for (const name of groups) {
      optionsHtml += `<option>${name}</option>`;
    }
    document.getElementById("group").innerHTML = optionsHtml;

    document.getElementById("status").textContent =
      `${allRows.length} inscriptions, ${groups.length} groupes`;

    updateStudentList();
  };

  const file = document.getElementById("file").files[0];
  reader.readAsArrayBuffer(file);
};

document.getElementById("group").onchange = () => {
  updateStudentList();
};

// texte = smiya b l'francais, value = smiya b l'arabe (an7tajoh arab name mn b3d)
function updateStudentList() {
  let studentsHtml = "";
  for (const s of getStudents()) {
    studentsHtml += `<option value="${s.Nom_Arabe} ${s.Prenom_arabe}">${s.Nom} ${s.Prenom}</option>`;
  }
  document.getElementById("dcStudent").innerHTML = studentsHtml;
}


// ---------Tabs Switch Mechanism 

function showTab(name) {
  document.getElementById("tabEff").hidden = (name !== "eff");
  document.getElementById("tabEfm").hidden = (name !== "efm");
  document.getElementById("tabDisc").hidden = (name !== "disc");
}
document.getElementById("btnTabEff").onclick = () => { showTab("eff"); };
document.getElementById("btnTabEfm").onclick = () => { showTab("efm"); };
document.getElementById("btnTabDisc").onclick = () => { showTab("disc"); };


// ---------- helper funcs

// stagiaires dyal groupe li mkhtar, bla doubles, mrettbin b smiya
// (nafs stagiaire kayji merra f kol groupe, so kanchoufo CEF deja zednah wla la)
function getStudents() {
  const group = document.getElementById("group").value;
  const students = [];
  const seen = [];

  for (const row of allRows) {
    const cef = String(row.MatriculeEtudiant);

    if (row.CodeDiplome === group && !seen.includes(cef)) {
      students.push(row);
      seen.push(cef);
    }
  }

  // localeCompare l ordre alphabetic
  students.sort((a, b) => (a.Nom + a.Prenom).localeCompare(b.Nom + b.Prenom));
  return students;
}

function getStudentsOrAlert() {
  const students = getStudents();
  if (students.length === 0) {
    alert("Chargez le fichier et choisissez un groupe.");
  }
  return students;
}

// example: "AG_INFO_TS_1A-Infographie (1A)-2025" -> "AG_INFO_TS"
function getCode(row) {
  const text = String(row.LibelleLong);
  const code = text.split("-")[0];
  return code.replace(/_\dA$/, "");
}

// example: "AG_INFO_TS_1A-Infographie (1A)-2025" -> "Infographie"
function getFiliere(row) {
  const text = String(row.LibelleLong);
  const name = text.substring(text.indexOf("-") + 1);
  return name.replace(/\s*\(\dA\).*$/, "");
}


// example: "AG_INFO_TS" ends with _TS -> "TECHNICIEN SPECIALISE"
function getLevel(row) {
  const code = getCode(row);
  if (/_TS(_|$)/.test(code)) { return "TECHNICIEN SPECIALISE"; }
  if (/_FQ(_|$)/.test(code)) { return "FORMATION QUALIFIANTE"; }
  if (/_S(_|$)/.test(code)) { return "SPECIALISATION"; }
  if (/_Q(_|$)/.test(code)) { return "QUALIFICATION"; }
  return "TECHNICIEN";
}

function shortDate(d) {
  return d.toLocaleDateString("fr-FR", { timeZone: "UTC" });
}

function longDate(d) {
  return d.toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric", timeZone: "UTC" });
}

// kanqraw <input type="date"> w kanrjo3 Date, ola null ila khawi (empty -> null)
// T12:00 bach l'jour may w9e3ch mochkil f date (js par default f date ka ybda m3a 12 dlil so redinah bda m3a 12pm instead)
function readDate(id) {
  const value = document.getElementById(id).value;
  if (value === "") {
    return null;
  }
  return new Date(value + "T12:00:00Z");
}


// --------= Documents 

document.getElementById("btnConvocation").onclick = () => {
  const students = getStudentsOrAlert();
  if (students.length === 0) {
    return;
  }

  const d = readDate("date");
  const dateText = (d !== null) ? longDate(d) : "";

  // convocation: page wa7da l kol stagiaire (one page per student)
  let html = "";

  for (const s of students) {
    html += `
      <div dir="rtl">
        <p dir="ltr"><b>Office de la Formation Professionnelle<br>et de la Promotion du Travail</b></p>
        <h2>استدعاء اجتياز امتحانات نهاية التكوين</h2>

        <table border="1" cellspacing="0" cellpadding="4" dir="ltr">
          <tr><td>CEF: ${s.MatriculeEtudiant}</td><td>FILIERE: ${getFiliere(s)}</td></tr>
          <tr><td>NOM: ${s.Nom}</td><td>ANNEE: ${document.getElementById("year").value}</td></tr>
          <tr><td>PRENOM: ${s.Prenom}</td><td>GROUPE: ${s.CodeDiplome}</td></tr>
          <tr><td>CNI: ${s.CIN}</td><td>NIVEAU: ${getLevel(s)}</td></tr>
        </table>

        <p>أنتم مدعوون لاجتياز امتحانات نهاية السنة برسم دورة ${document.getElementById("session").value}
        حسب البرنامج المحدد أسفله، لذا عليكم الحضور إلى مقر الامتحان 30 دقيقة قبل بداية الاختبارات، مزودين بهذه الدعوة وبطاقة التعريف الوطنية.</p>

        <h3>برنامج الامتحان</h3>
        <table border="1" cellspacing="0" cellpadding="4" dir="ltr">
          <tr><th>Date</th><th>Horaires</th><th>Salle / Atelier</th><th>Etablissement</th><th>Épreuve</th></tr>
          <tr>
            <td>${dateText}</td>
            <td>${document.getElementById("heure").value}</td>
            <td>${document.getElementById("salle").value}</td>
            <td>CMFP AGADIR</td>
            <td>${document.getElementById("epreuve").value}</td>
          </tr>
        </table>

        <p>التدابير والإجراءات المتعلقة بالامتحانات منصوص عليها في القانون الداخلي وخاصة الفصل السادس (F) منه:</p>
        <ul>
          <li>يمنع منعا كليا إحضار إلى فضاء إجراء الامتحانات أي وسيلة للتواصل (هاتف أو أية واسطة إلكترونية أو غيرها....).</li>
          <li>إجبارية اللباس الموحد / البطاقة التعريفية للمتدرب / الهندام اللائق.</li>
          <li>وجوب إحضار جميع أدوات العمل حيث يمنع تبادلها أثناء الامتحان.</li>
          <li>كل محاولة للغش تعرض صاحبها للإقصاء والحرمان من الامتحان لمدة ثلاث سنوات.</li>
        </ul>

        <p><b>الإدارة</b></p>
        <hr>
      </div>`;
  }

  document.getElementById("output").innerHTML = html;
};

document.getElementById("btnPresence").onclick = () => {
  const students = getStudentsOrAlert();
  if (students.length === 0) {
    return;
  }

  // stagiaire wa7ed kafi bach njibo les info (filiere / groupe / niveau)
  const studentForInfo = students[0];
  const d = readDate("date");
  const dateText = (d !== null) ? shortDate(d) : "……";
  const year = document.getElementById("year").value;

  // <tr> wa7da l kol stagiaire 
  let rows = "";
  for (let i = 0; i < students.length; i++) {
    const s = students[i];
    rows += `
      <tr>
        <td>${i + 1}</td>
        <td>${s.MatriculeEtudiant}</td>
        <td>${s.Nom} ${s.Prenom}</td>
        <td></td><td></td><td></td>
      </tr>`;
  }

  document.getElementById("output").innerHTML = `
    <p><b>OFPPT - RENTREE ${year.replace("-", "/")}</b> | DRSM - CMFP Agadir</p>
    <h2>Feuille de présence des stagiaires – Examens de Fin de Formation</h2>
    <p>Filière: ${getFiliere(studentForInfo)} | Groupe: ${studentForInfo.CodeDiplome} | Niveau: ${getLevel(studentForInfo)}<br>
      Épreuve: ${document.getElementById("epreuve").value} | Date: ${dateText}
      | Heure: ${document.getElementById("heure").value} | Salle: ${document.getElementById("salle").value}
      | Session: ${document.getElementById("session").value}</p>
    <table border="1" cellspacing="0" cellpadding="4">
      <tr><th>N°</th><th>N° inscription</th><th>Nom et Prénom</th>
          <th>Émargement début</th><th>Émargement fin</th><th>Observations</th></tr>
      ${rows}
    </table>`;
};

document.getElementById("btnListe").onclick = () => {
  const students = getStudentsOrAlert();
  if (students.length === 0) {
    return;
  }

  // same as above
  const studentForInfo = students[0];
  const d = readDate("date");
  const dateText = (d !== null) ? shortDate(d) : "……";
  const year = document.getElementById("year").value;

  // <tr> wa7da l kol stagiaire same as above again
  let rows = "";
  for (let i = 0; i < students.length; i++) {
    const s = students[i];
    rows += `
      <tr>
        <td>${i + 1}</td>
        <td>${s.MatriculeEtudiant}</td>
        <td>${s.Nom}</td>
        <td>${s.Prenom}</td>
      </tr>`;
  }

  document.getElementById("output").innerHTML = `
    <p><b>OFPPT - RENTREE ${year.replace("-", "/")}</b> | DRSM - CMFP Agadir</p>
    <h2>EXAMEN DE FIN DE FORMATION</h2>
    <p>Filière: ${getFiliere(studentForInfo)} | Groupe: ${studentForInfo.CodeDiplome} | Niveau: ${getLevel(studentForInfo)}<br>
      Épreuve: ${document.getElementById("epreuve").value} | Date: ${dateText}
      | Heure: ${document.getElementById("heure").value} | Salle: ${document.getElementById("salle").value}
      | Session: ${document.getElementById("session").value}</p>
    <table border="1" cellspacing="0" cellpadding="4">
      <tr><th>N° Ordre</th><th>CEF</th><th>Nom</th><th>Prénom</th></tr>
      ${rows}
    </table>`;
};

document.getElementById("btnEfm").onclick = () => {
  const students = getStudentsOrAlert();
  if (students.length === 0) {
    return;
  }

  // again same as above
  const studentForInfo = students[0];
  const year = document.getElementById("year").value;

  // les champs khawyin = no9at for manual typing
  const moduleName = document.getElementById("emModule").value || "……";
  const horaire = document.getElementById("emHoraire").value || "……";

  const d = readDate("emDate");
  const dateText = (d !== null) ? shortDate(d) : "……";

  // <tr> wa7da l kol stagiaire again
  let rows = "";
  for (let i = 0; i < students.length; i++) {
    const s = students[i];
    rows += `
      <tr>
        <td>${i + 1}</td>
        <td>${s.MatriculeEtudiant}</td>
        <td>${s.Nom}</td>
        <td>${s.Prenom}</td>
        <td></td><td></td><td></td>
      </tr>`;
  }

  document.getElementById("output").innerHTML = `
    <p><b>OFPPT - RENTREE ${year.replace("-", "/")}</b> | DRSM - CMFP Agadir</p>
    <h2>${document.getElementById("emType").value}</h2>
    <p><b>LISTE DE PRÉSENCE DES PARTICIPANTS</b><br>
      Niveau: ${getLevel(studentForInfo)} | Filière: ${getFiliere(studentForInfo)} (${studentForInfo.CodeDiplome})<br>
      Module: ${moduleName} | Date: ${dateText} | Horaire: ${horaire}<br>
      Effectif stagiaires: ${students.length} | Effectif présents: ……</p>
    <table border="1" cellspacing="0" cellpadding="4">
      <tr><th>N°</th><th>CEF</th><th>Nom</th><th>Prénom</th>
          <th>Signature avant</th><th>Signature après</th><th>Obs.</th></tr>
      ${rows}
    </table>`;
};


// ---------- printing 

document.getElementById("btnPrint").onclick = () => {

  const content = document.getElementById("output").innerHTML;

  if (content === "") {
    alert("Generez d'abord un document.");
    return;
  }

  // ka n7elo fenetre jdida fiha ghir document li kayna f div d output, w kanprintiw dik new window
  const printWindow = window.open("", "_blank");
  printWindow.document.write(`<html><head><meta charset="utf-8"><title>Impression</title></head><body>${content}</body></html>`);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
};
