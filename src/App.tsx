import { useState } from "react";
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
  "ינו-26","פבר-26","מרץ-26","אפר-26","מאי-26","יונ-26",
  "יול-26","אוג-26","ספט-26","אוק-26","נוב-26","דצמ-26",
];

// =====================
// יעד חכם
// =====================
function computeSmartTarget(history: MonthRecord[]) {
  const healthy = history.filter(
    (m) => m.status !== "negative" && m.K > 0
  );

  if (healthy.length < 2) return 0;

  const avg =
    healthy.reduce((sum, m) => sum + m.K, 0) / healthy.length;

  return Math.round(avg * 0.9 * 100) / 100;
}

function App() {
  const [month, setMonth] = useState("");
  const [reg, setReg] = useState(0);
  const [h125, setH125] = useState(0);
  const [h150, setH150] = useState(0);
  const [gross, setGross] = useState(0);
  const [ded, setDed] = useState(0);

  const [history, setHistory] = useState<MonthRecord[]>([]);
  const [targetHourlyNet, setTargetHourlyNet] = useState(0);

  const weightedHours = reg + h125 * 1.25 + h150 * 1.5;
  const net = gross - ded;
  const K = weightedHours > 0 ? net / weightedHours : 0;

  // =====================
  // טעינת אקסל
  // =====================
  function handleFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (evt) => {
      const data = evt.target?.result;
      if (!data) return;

      const workbook = XLSX.read(data, { type: "binary" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

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
        let S: number | undefined;

        if (imported.length > 0) {
          const prev = imported[imported.length - 1];
          S = (net - prev.net) / (hours - prev.hours);

          if (isFinite(S)) {
            if (S < prev.K - 5) status = "negative";
            else if (S > prev.K + 5) status = "positive";
          }
        }

        imported.push({ month, hours, net, K, S, status });
      });

      setHistory(imported);
      setTargetHourlyNet(computeSmartTarget(imported));
    };

    reader.readAsBinaryString(file);
  }

  // =====================
  // הוספה ידנית
  // =====================
  function addMonth() {
    if (!month || weightedHours <= 0) return;

    let status: Status = "neutral";
    let S: number | undefined;

    if (history.length > 0) {
      const prev = history[history.length - 1];
      S = (net - prev.net) / (weightedHours - prev.hours);

      if (isFinite(S)) {
        if (targetHourlyNet > 0 && S < targetHourlyNet) {
          status = "negative";
        } else if (S > prev.K + 5) {
          status = "positive";
        } else if (S < prev.K - 5) {
          status = "negative";
        }
      }
    }

    setHistory([...history, { month, hours: weightedHours, net, K, S, status }]);

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
      <h1>כמה באמת שווה לך כל שעה?</h1>

      {/* הסבר על השירות – לפני הכל */}
      <p className="intro">
        הכלי הזה עוזר להבין האם משתלם לך לעבוד יותר שעות,
        או שהשעות הנוספות כבר לא משתלמות בגלל מיסוי וניכויים.
        <br />
        טוענים נתוני שכר (או מזינים ידנית), והמערכת מחשבת
        כמה באמת הרווחת על כל שעה נוספת – וממליצה אם כדאי
        <strong> להוסיף שעות</strong>, <strong>להישאר כמו שאתה</strong>,
        או <strong>להוריד שעות</strong>.
      </p>

      {/* יעד */}
      <div className="card">
        <label>
          🎯 יעד נטו מינימלי לשעה (מחושב אוטומטית)
          <input
            type="number"
            value={targetHourlyNet}
            onChange={(e) => setTargetHourlyNet(+e.target.value)}
          />
        </label>
        <p className="note">
          אם שעה נוספת שווה פחות מהיעד – זו נחשבת פגיעה.
        </p>
      </div>

      {/* מקרא אקסל */}
      <div className="legend">
        <h3>📊 מבנה קובץ האקסל</h3>
        <p>
          הקובץ חייב להכיל שורת כותרות ועמודות A–F בלבד.
          עמודות נוספות לא משפיעות.
        </p>
      </div>

      {/* טעינה */}
      <div className="card">
        <label>
          📂 טעינת קובץ אקסל
          <input type="file" accept=".xlsx,.xls" onChange={handleFile} />
        </label>
      </div>

      {/* הזנה ידנית */}
      <div className="card">
        <label>
          חודש
          <select value={month} onChange={(e) => setMonth(e.target.value)}>
            <option value="">בחר חודש</option>
            {MONTH_OPTIONS.map((m) => (
              <option key={m} value={m}>{m}</option>
            ))}
          </select>
        </label>

        <label>שעות רגילות
          <input type="number" value={reg} onChange={(e) => setReg(+e.target.value)} />
        </label>

        <label>שעות 125%
          <input type="number" value={h125} onChange={(e) => setH125(+e.target.value)} />
        </label>

        <label>שעות 150%
          <input type="number" value={h150} onChange={(e) => setH150(+e.target.value)} />
        </label>

        <label>שכר ברוטו (₪)
          <input type="number" value={gross} onChange={(e) => setGross(+e.target.value)} />
        </label>

        <label>ניכויים (₪)
          <input type="number" value={ded} onChange={(e) => setDed(+e.target.value)} />
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

      {/* הסבר תחתון – איך מחושבת ההמלצה */}
      <div className="explanation">
        <h3>איך מחושבת ההמלצה?</h3>
        <ul>
          <li><strong>➕ להוסיף</strong> – השעה הנוספת משתלמת יותר מהממוצע</li>
          <li><strong>➗ להשאיר</strong> – אין שינוי מהותי בכדאיות</li>
          <li><strong>➖ להוריד</strong> – השעה הנוספת שווה פחות (בד״כ בגלל מיסוי)</li>
        </ul>
        <p className="note">
          ההשוואה מתבצעת מול החודש הקודם. חודשים עם תיקונים רטרואקטיביים
          עשויים להיראות חריגים.
        </p>
      </div>
    </div>
  );
}

export default App;
