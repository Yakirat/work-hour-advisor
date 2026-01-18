import { useEffect, useState } from "react"; // ⭐ NEW: useEffect
import * as XLSX from "xlsx";
import "./App.css";

type Status = "positive" | "neutral" | "negative";

type MonthRecord = {
  month: string;
  hours: number;
  net: number;
  K: number;
  S?: number;
  status: Status;
};

const MONTH_OPTIONS = [
  "ינו-26",
  "פבר-26",
  "מרץ-26",
  "אפר-26",
  "מאי-26",
  "יונ-26",
  "יול-26",
  "אוג-26",
  "ספט-26",
  "אוק-26",
  "נוב-26",
  "דצמ-26",
];

// ⭐ NEW: חישוב יעד ברירת מחדל חכם
function calcSmartTargetHours(history: MonthRecord[]) {
  const healthy = history.filter(
    (m) => m.status !== "negative" && m.hours > 0
  );
  if (healthy.length === 0) return 0;

  const avg =
    healthy.reduce((sum, m) => sum + m.hours, 0) / healthy.length;

  return Math.round(avg * 0.9);
}

function App() {
  const [month, setMonth] = useState("");
  const [reg, setReg] = useState(0);
  const [h125, setH125] = useState(0);
  const [h150, setH150] = useState(0);
  const [gross, setGross] = useState(0);
  const [ded, setDed] = useState(0);

  const [history, setHistory] = useState<MonthRecord[]>([]);

  // ⭐ NEW: יעד חודשי + דגל שינוי ידני
  const [targetHours, setTargetHours] = useState(0);
  const [isTargetManual, setIsTargetManual] = useState(false);

  const weightedHours = reg + h125 * 1.25 + h150 * 1.5;
  const net = gross - ded;
  const K = weightedHours > 0 ? net / weightedHours : 0;

  // ⭐ NEW: עדכון אוטומטי של היעד אחרי טעינת אקסל / הוספת חודש
  useEffect(() => {
    if (isTargetManual) return;

    const smart = calcSmartTargetHours(history);
    if (smart > 0) setTargetHours(smart);
  }, [history, isTargetManual]);

    // ⭐ NEW: חישוב הקשר ליעד עבור החודש האחרון
  const lastMonth =
    history.length > 0 ? history[history.length - 1] : null;

  const diffFromTarget =
    lastMonth && targetHours > 0
      ? Math.round(lastMonth.hours - targetHours)
      : null;

  // =====================
  // טעינת קובץ אקסל
  // =====================
  function handleFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (evt) => {
      const data = evt.target?.result;
      if (!data) return;

      const workbook = XLSX.read(data, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const rows: any[][] = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
      });

      const dataRows = rows.slice(1);
      const imported: MonthRecord[] = [];

      dataRows.forEach((row) => {
        if (!row[0]) return;

        const month = String(row[0]);
        const reg = Number(row[1] || 0);
        const h125 = Number(row[2] || 0);
        const h150 = Number(row[3] || 0);
        const gross = Number(row[4] || 0);
        const ded = Number(row[5] || 0);

        const hours = reg + h125 * 1.25 + h150 * 1.5;
        const net = gross - ded;
        const K = hours > 0 ? net / hours : 0;

        let status: Status = "neutral";
        let S: number | undefined = undefined;

        if (imported.length > 0) {
          const prev = imported[imported.length - 1];
          S = (net - prev.net) / (hours - prev.hours);

          if (S < prev.K - 5) status = "negative";
          else if (S > prev.K + 5) status = "positive";
        }

        imported.push({ month, hours, net, K, S, status });
      });

      setHistory(imported);
    };

    reader.readAsBinaryString(file);
  }

  // =====================
  // הוספה ידנית של חודש
  // =====================
  function addMonth() {
    if (!month || weightedHours <= 0) return;

    let status: Status = "neutral";
    let S: number | undefined = undefined;

    if (history.length > 0) {
      const prev = history[history.length - 1];
      S = (net - prev.net) / (weightedHours - prev.hours);

      if (S < prev.K - 5) status = "negative";
      else if (S > prev.K + 5) status = "positive";
    }

    const record: MonthRecord = {
      month,
      hours: weightedHours,
      net,
      K,
      S,
      status,
    };

    setHistory([...history, record]);

    setMonth("");
    setReg(0);
    setH125(0);
    setH150(0);
    setGross(0);
    setDed(0);
  }

  // =====================
  // UI
  // =====================
  return (
    <div className="container">
      <h1>💰  מחשבון לניתוח אפקטיביות ומיסוי שעות עבודה נוספות  💰</h1>

    <p className="intro">
      הכלי הזה עוזר להבין האם משתלם לך לעבוד יותר שעות,
      או שדווקא השעות הנוספות כבר לא משתלמות בגלל מיסוי וניכויים.
      <br />
      טוענים נתוני שכר (או מזינים ידנית), והמערכת מחשבת כמה באמת
      הרווחת על כל שעה נוספת – וממליצה אם כדאי
      <strong> להוסיף שעות</strong>, <strong>להישאר כמו שאתה</strong>,
      או <strong>להוריד שעות</strong>.
    </p>

      <div className="legend">
        <h3>📊 מבנה קובץ האקסל</h3>

        <p>
          כדי לטעון נתונים למערכת, יש להכין קובץ אקסל עם
          <strong> שורה ראשונה ככותרות </strong>
          והעמודות הבאות (בסדר הזה):
        </p>

        <table className="legend-table">
          <thead>
            <tr>
              <th>עמודה</th>
              <th>שם</th>
              <th>חובה</th>
              <th>הערה</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>A</td>
              <td>חודש</td>
              <td>כן</td>
              <td>לדוגמה: אוג-25</td>
            </tr>
            <tr>
              <td>B</td>
              <td>שעות רגילות</td>
              <td>כן</td>
              <td>מספר</td>
            </tr>
            <tr>
              <td>C</td>
              <td>שעות נוספות 125%</td>
              <td>כן</td>
              <td>מספר</td>
            </tr>
            <tr>
              <td>D</td>
              <td>שעות נוספות 150%</td>
              <td>כן</td>
              <td>מספר</td>
            </tr>
            <tr>
              <td>E</td>
              <td>שכר ברוטו</td>
              <td>כן</td>
              <td>בש״ח</td>
            </tr>
            <tr>
              <td>F</td>
              <td>ניכויים</td>
              <td>כן</td>
              <td>סה״כ ניכויים</td>
            </tr>
            <tr>
              <td>G–T</td>
              <td>כל דבר אחר</td>
              <td>לא</td>
              <td>המערכת מתעלמת</td>
            </tr>
          </tbody>
        </table>

        <p className="note">
          אין צורך בעמודות נוספות. נתונים שאינם בעמודות A–F לא משפיעים על החישוב.
          נתוני הברוטו לא יכללו תוספות או בונוסים חד פעמיים.
        </p>
      </div>
      {/* טעינת אקסל */}
      <div className="card">
        <label>
          📂 טעינת קובץ אקסל
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFile}
          />
        </label>
      </div>

      {/* ⭐ NEW: יעד חודשי */}
      <div className="card">
      <h3>
        🎯 יעד שעות חודשי
        <span className="tooltip">
          ℹ️
          <span className="tooltip-text">
            היעד מחושב לפי החודשים הבריאים שלך:
            ממוצע שעות → 90% → עיגול.
            משמש כהמלצה ונקודת ייחוס בלבד.
          </span>
        </span>
      </h3>

        <p className="note">
          זה מספר השעות שבו, לפי ההיסטוריה שלך, אתה מרוויח בצורה
          האופטימלית – בלי שהשעות הנוספות נשחקות ממס וניכויים.
        </p>

        <label>
          יעד (שעות)
          <input
            type="number"
            value={targetHours}
            onChange={(e) => {
              setTargetHours(+e.target.value);
              setIsTargetManual(true);
            }}
          />
        </label>

        <button
          onClick={() => {
            setTargetHours(calcSmartTargetHours(history));
            setIsTargetManual(false);
          }}
          disabled={history.length === 0}
        >
          ♻️ חזור ליעד חכם
        </button>

        {/* ⭐ NEW: הסבר הקשר ליעד */}
        {lastMonth && targetHours > 0 && diffFromTarget !== null && (
          <div className="note" style={{ marginTop: 12 }}>
            <p>
              📅 <strong>החודש האחרון:</strong>{" "}
              {lastMonth.hours.toFixed(1)} שעות
            </p>

            {diffFromTarget > 0 && (
              <p>
                ⚠️ <strong>חריגה:</strong> {diffFromTarget} שעות מעל היעד
              </p>
            )}

            {diffFromTarget < 0 && (
              <p>
                📉 <strong>מתחת ליעד:</strong>{" "}
                {Math.abs(diffFromTarget)} שעות
              </p>
            )}

            {diffFromTarget === 0 && (
              <p>🟢 אתה בדיוק על היעד</p>
            )}

            <p>
              🧠 <strong>מסקנה:</strong>{" "}
              {diffFromTarget > 5 &&
                "עבדת מעבר לאזור האופטימלי שלך. סביר שהשעות הנוספות פחות משתלמות."}
              {diffFromTarget < -5 &&
                "יש לך מרווח להוסיף שעות בלי פגיעה משמעותית ברווח לשעה."}
              {Math.abs(diffFromTarget) <= 5 &&
                "אתה עובד באזור האופטימלי שלך. אין סיבה לשנות כרגע."}
            </p>
          </div>
        )}
      </div>


      {/* הזנה ידנית */}
      <div className="card">
        <label>
          חודש
          <select value={month} onChange={(e) => setMonth(e.target.value)}>
            <option value="">בחר חודש</option>
            {MONTH_OPTIONS.map((m) => (
              <option key={m} value={m}>
                {m}
              </option>
            ))}
          </select>
        </label>

        <label>
          שעות רגילות
          <input
            type="number"
            value={reg}
            onChange={(e) => setReg(+e.target.value)}
          />
        </label>

        <label>
          שעות 125%
          <input
            type="number"
            value={h125}
            onChange={(e) => setH125(+e.target.value)}
          />
        </label>

        <label>
          שעות 150%
          <input
            type="number"
            value={h150}
            onChange={(e) => setH150(+e.target.value)}
          />
        </label>

        <label>
          שכר ברוטו (₪)
          <input
            type="number"
            value={gross}
            onChange={(e) => setGross(+e.target.value)}
          />
        </label>

        <label>
          ניכויים סה״כ (₪)
          <input
            type="number"
            value={ded}
            onChange={(e) => setDed(+e.target.value)}
          />
        </label>

        <button onClick={addMonth}>➕ הוסף חודש</button>
      </div>

      {/* טבלה */}
      {history.length > 0 && (
        <table>
          <thead>
            <tr>
              <th>חודש</th>
              <th>שעות</th>
              <th>נטו/שעה</th>
              <th>נטו שולי</th>
              <th>מצב</th>
              <th>המלצה</th>
            </tr>
          </thead>
          <tbody>
            {history.map((m, i) => (
              <tr key={i} className={m.status}>
                <td>{m.month}</td>
                <td>{m.hours.toFixed(1)}</td>
                <td>{m.K.toFixed(2)}</td>
                <td>{m.S !== undefined ? m.S.toFixed(2) : "-"}</td>
                <td>{m.status}</td>
                <td>
                  {m.status === "positive" && "➕ להוסיף"}
                  {m.status === "neutral" && "➗ להשאיר"}
                  {m.status === "negative" && "➖ להוריד"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
          <div className="explanation">
      <h3>איך מחושבת ההמלצה?</h3>

      <p>
        ההמלצה מבוססת על השוואה בין <strong>הרווח נטו לשעה</strong> לבין
        <strong> הרווח מהשעות הנוספות</strong> שנוספו החודש.
      </p>

      <ul>
        <li>
          <strong>➕ להוסיף שעות</strong> – אם כל שעה נוספת הכניסה לך יותר
          מהשכר הממוצע לשעה.
        </li>
        <li>
          <strong>➗ להשאיר</strong> – אם אין שינוי מהותי ברווח לשעה.
        </li>
        <li>
          <strong>➖ להוריד שעות</strong> – אם השעות הנוספות שוות פחות
          משמעותית, לרוב בגלל מיסוי גבוה יותר.
        </li>
      </ul>

      <p className="note">
        ההשוואה מתבצעת מול החודש הקודם בלבד, ולכן חודשים עם תיקונים רטרואקטיביים
        עשויים להיראות חריגים.
      </p>
      <p className="note">
        הכלי מיועד למתן אינדיקציה כללית בלבד ואינו מהווה ייעוץ פיננסי או מס.
      </p>
      <p className="note">
        המערכת לא שומרת נתונים אישיים או פיננסיים כלשהם. רענון הדף יוביל לאיפוס
        כל הנתונים שהוזנו.
      </p>
    </div>
    </div>
  );
}

export default App;
