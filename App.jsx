import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import {
  Search, UploadCloud, ChevronDown, ChevronLeft, X, Waves,
  Trash2, Info, RefreshCw, Database, ShieldCheck
} from "lucide-react";
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || "";
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || "";
const supabase = SUPABASE_URL && SUPABASE_ANON_KEY
  ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
  : null;

const T = {
  ink: "#0B2530", deep: "#0B3D54", water: "#1C7293",
  aqua: "#EAF6F8", card: "#FFFFFF", gold: "#D4A72C",
  silver: "#9AA5AB", bronze: "#B5651D", line: "#D8E7EA"
};

function toSeconds(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number") {
    // Excel stores a time-of-day as a fraction of a day.
    if (value > 0 && value < 1) return value * 86400;
    return value;
  }
  const s = String(value).trim().replace(",", ".");
  let m = s.match(/^(\d+):(\d+(?:\.\d+)?)$/);
  if (m) return Number(m[1]) * 60 + Number(m[2]);
  m = s.match(/^(\d+):(\d+):(\d+(?:\.\d+)?)$/);
  if (m) return Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3]);
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function fmtSeconds(sec) {
  if (sec == null) return "—";
  const hundredths = Math.round(sec * 100) / 100;
  if (hundredths >= 60) {
    const m = Math.floor(hundredths / 60);
    const s = (hundredths - m * 60).toFixed(2).padStart(5, "0");
    return `${m}:${s}`;
  }
  return hundredths.toFixed(2);
}

function clean(v) {
  return v == null ? "" : String(v).trim();
}

function scanMeta(rows) {
  let title1 = "", title2 = "", dateLabel = "", pool = "", city = "";
  const limit = Math.min(rows.length, 15);
  for (let i = 0; i < limit; i++) {
    for (const cell of rows[i] || []) {
      const text = clean(cell);
      if (!text) continue;
      if (!title1 && /чемпионат|первенств|турнир|кубок|спартакиад/i.test(text)) title1 = text;
      if (!title2 && title1 && text !== title1 && i <= 2 && !/тел|организ|www|@/i.test(text)) title2 = text;
      if (!dateLabel && (/\d{4}\s*г/i.test(text) || /январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр/i.test(text))) dateLabel = text;
      if (!pool && /бассейн|п\/б|\b\d+\s*м\b/i.test(text)) pool = text;
      if (!city && /^г\.\s*/i.test(text)) city = text;
    }
  }
  return {
    title: [title1, title2].filter(Boolean).join(" ").replace(/\s+/g, " ").trim() || "Соревнования",
    dateLabel,
    pool: [city, pool].filter(Boolean).join(", ")
  };
}

function parseSheetRows(rows, dayLabel, out) {
  const headerIdx = rows.findIndex(r => r && clean(r[0]) === "№");
  if (headerIdx < 0) return;
  let currentEvent = null;
  let currentAgeGroup = null;
  let emptyStreak = 0;

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const [place, name, byear, region, result, category, zachet, coach] = row;
    const allEmpty = row.slice(0, 8).every(v => clean(v) === "");
    if (allEmpty) {
      emptyStreak++;
      if (emptyStreak > 30) break;
      continue;
    }
    emptyStreak = 0;

    if (!clean(name)) {
      const text = clean(place);
      if (!text) continue;
      if (/г\.р|года рождения|\d{4}\s*[-–]\s*\d{4}/i.test(text)) currentAgeGroup = text;
      else currentEvent = text;
      continue;
    }

    out.push({
      day: dayLabel,
      event: currentEvent || "—",
      ageGroup: currentAgeGroup || "",
      place: Number.isFinite(Number(place)) ? Number(place) : null,
      name: clean(name),
      birthYear: byear ?? null,
      region: clean(region) || null,
      result: result ?? null,
      category: clean(category) || null,
      zachet: clean(zachet) || null,
      coach: clean(coach) || null
    });
  }
}

function parseWorkbook(workbook) {
  let meta = null;
  const records = [];
  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    if (!meta) meta = scanMeta(rows);
    let dayLabel = sheetName;
    for (let i = 0; i < Math.min(rows.length, 12); i++) {
      const text = clean(rows[i]?.[0]);
      if (/день/i.test(text)) { dayLabel = text; break; }
    }
    parseSheetRows(rows, dayLabel, records);
  }
  return { ...(meta || { title: "Соревнования", dateLabel: "", pool: "" }), records };
}

function PlaceBadge({ place }) {
  if (!place) return <span className="place-empty">—</span>;
  const colors = {1: T.gold, 2: T.silver, 3: T.bronze};
  return <span className="place-badge" style={{background: colors[place] || T.deep}}>{place}</span>;
}

function Rope() {
  return <div className="rope"><div/><span>●●●●●●●●●●●●●●●●●●●●</span><div/></div>;
}

function SwimmerCard({ name, results, meetsById }) {
  const byEvent = useMemo(() => {
    const map = {};
    results.forEach(r => {
      const key = r.event || "—";
      (map[key] ||= []).push(r);
    });
    Object.values(map).forEach(arr => arr.sort((a,b) =>
      (meetsById[a.meetId]?.uploadedAt || 0) - (meetsById[b.meetId]?.uploadedAt || 0)
    ));
    return map;
  }, [results, meetsById]);

  return <div className="swimmer-card">
    <div className="swimmer-head">
      <div className="swimmer-name">{name}</div>
      <div>{results.length} {results.length === 1 ? "результат" : "результатов"} · {Object.keys(byEvent).length} дисципл.</div>
    </div>
    {Object.entries(byEvent).map(([event, arr]) => {
      const valid = arr.map(r => ({...r, sec: toSeconds(r.result)})).filter(r => r.sec != null);
      const best = valid.length ? Math.min(...valid.map(r => r.sec)) : null;
      const chartData = valid.map(r => ({
        date: meetsById[r.meetId]?.dateLabel || "",
        sec: r.sec
      }));
      return <div className="event-section" key={event}>
        <div className="event-title">
          <b>{event}</b>
          {best != null && <span>Личный рекорд: <strong>{fmtSeconds(best)}</strong></span>}
        </div>
        {arr.map((r,i) => {
          const meet = meetsById[r.meetId];
          return <div className="result-line" key={i}>
            <span>{meet?.title || "Соревнования"}{meet?.dateLabel ? `, ${meet.dateLabel}` : ""}</span>
            <span className="result-right"><PlaceBadge place={r.place}/><strong className={r.sec === best ? "best" : ""}>{r.result != null ? String(r.result) : "—"}</strong></span>
          </div>
        })}
        {chartData.length >= 2 && <div className="chart">
          <ResponsiveContainer width="100%" height={100}>
            <LineChart data={chartData}>
              <XAxis dataKey="date" tick={{fontSize:9}} />
              <YAxis hide domain={["auto","auto"]}/>
              <Tooltip formatter={v => fmtSeconds(v)} />
              <Line type="monotone" dataKey="sec" stroke={T.water} strokeWidth={2} dot={{r:3}}/>
            </LineChart>
          </ResponsiveContainer>
          <small>меньше время — лучше результат</small>
        </div>}
      </div>
    })}
  </div>;
}

function MeetBrowser({ meets, onDelete }) {
  const [openMeet, setOpenMeet] = useState(null);
  const [openEvent, setOpenEvent] = useState(null);
  if (!meets.length) return null;
  const meet = meets.find(m => m.id === openMeet);

  if (meet) {
    const events = {};
    meet.records.forEach(r => {
      const key = `${r.day || ""}||${r.event || "—"}`;
      (events[key] ||= {day:r.day,event:r.event,rows:[]}).rows.push(r);
    });
    if (openEvent && events[openEvent]) {
      const ev = events[openEvent];
      return <div>
        <button className="back" onClick={() => setOpenEvent(null)}><ChevronLeft size={16}/> {ev.event}</button>
        <div className="table-card">
          {ev.rows.slice().sort((a,b)=>(a.place??999)-(b.place??999)).map((r,i)=>
            <div className="result-row" key={i}><PlaceBadge place={r.place}/><div className="row-name"><b>{r.name}</b><small>{r.region || ""}</small></div><strong>{r.result != null ? String(r.result) : "—"}</strong></div>
          )}
        </div>
      </div>;
    }
    return <div>
      <button className="back" onClick={() => {setOpenMeet(null);setOpenEvent(null)}}><ChevronLeft size={16}/> Все соревнования</button>
      <div className="meet-title">{meet.title}</div>
      <div className="muted">{meet.dateLabel} {meet.pool ? `· ${meet.pool}` : ""}</div>
      <div className="event-list">
        {Object.entries(events).map(([k,e]) =>
          <button className="event-button" key={k} onClick={() => setOpenEvent(k)}>
            <span><b>{e.event}</b><small>{e.day} · {e.rows.length} участников</small></span>
            <ChevronDown size={17} style={{transform:"rotate(-90deg)"}}/>
          </button>
        )}
      </div>
    </div>;
  }

  return <div className="meet-list">{meets.map(m =>
    <div className="meet-card" key={m.id}>
      <button className="meet-open" onClick={() => setOpenMeet(m.id)}>
        <b>{m.title}</b><small>{m.dateLabel} {m.pool ? `· ${m.pool}` : ""} · {m.records.length} результатов</small>
      </button>
      <button className="delete" title="Удалить протокол" onClick={() => onDelete(m.id)}><Trash2 size={16}/></button>
    </div>
  )}</div>;
}

async function loadMeets() {
  if (!supabase) throw new Error("Supabase не настроен");
  const {data, error} = await supabase.from("meets").select("*").order("uploaded_at",{ascending:false});
  if (error) throw error;
  return (data || []).map(m => ({...m, id:m.id, title:m.title, dateLabel:m.date_label, pool:m.pool, records:m.records || [], uploadedAt:new Date(m.uploaded_at).getTime()}));
}

async function saveMeet(parsed) {
  if (!supabase) throw new Error("Supabase не настроен");
  const {data,error} = await supabase.from("meets").insert({
    title: parsed.title, date_label: parsed.dateLabel, pool: parsed.pool,
    records: parsed.records
  }).select().single();
  if (error) throw error;
  return {...data, id:data.id, dateLabel:data.date_label, uploadedAt:new Date(data.uploaded_at).getTime()};
}

async function deleteMeet(id) {
  if (!supabase) throw new Error("Supabase не настроен");
  const {error} = await supabase.from("meets").delete().eq("id", id);
  if (error) throw error;
}

export default function App() {
  const [meets,setMeets] = useState([]);
  const [query,setQuery] = useState("");
  const [tab,setTab] = useState("search");
  const [loading,setLoading] = useState(true);
  const [message,setMessage] = useState("");
  const [error,setError] = useState("");

  const refresh = async () => {
    setLoading(true); setError("");
    try { setMeets(await loadMeets()); }
    catch(e) { setError(e.message || "Не удалось загрузить данные"); }
    finally { setLoading(false); }
  };

  useEffect(() => { refresh(); }, []);

  const meetsById = useMemo(() => Object.fromEntries(meets.map(m => [m.id,m])), [meets]);
  const records = useMemo(() => meets.flatMap(m => (m.records || []).map(r => ({...r,meetId:m.id}))), [meets]);

  const grouped = useMemo(() => {
    const tokens = query.trim().toLowerCase().split(/\s+/).filter(Boolean);
    const map = {};
    if (tokens.join("").length < 2) return map;
    records.forEach(r => {
      const name = clean(r.name);
      if (name && tokens.every(t => name.toLowerCase().includes(t))) (map[name] ||= []).push(r);
    });
    return map;
  }, [query,records]);

  const handleFile = async e => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setMessage("Загружаю протокол…"); setError("");
    try {
      const wb = XLSX.read(await file.arrayBuffer(), {type:"array",cellDates:false,raw:true});
      const parsed = parseWorkbook(wb);
      if (!parsed.records.length) throw new Error("В файле не найдены строки с результатами.");
      const saved = await saveMeet(parsed);
      setMeets(prev => [saved,...prev]);
      setMessage(`Загружено: ${parsed.title} · ${parsed.records.length} результатов`);
    } catch(e) { setError(e.message || "Не удалось прочитать файл"); setMessage(""); }
  };

  const handleDelete = async id => {
    if (!confirm("Удалить этот протокол?")) return;
    try { await deleteMeet(id); setMeets(prev => prev.filter(m => m.id !== id)); }
    catch(e) { setError(e.message || "Не удалось удалить протокол"); }
  };

  return <div className="app">
    <header>
      <div className="brand"><Waves size={23}/><span>ДОРОЖКА</span></div>
      <div className="subtitle">результаты заплывов твоего пловца</div>
      <div className="tabs">
        <button className={tab==="search"?"active":""} onClick={()=>setTab("search")}>Найти пловца</button>
        <button className={tab==="browse"?"active":""} onClick={()=>setTab("browse")}>Соревнования ({meets.length})</button>
      </div>
    </header>

    <main>
      {!supabase && <div className="warning"><Database size={18}/><div><b>Supabase пока не подключён.</b><br/>Создай проект и добавь VITE_SUPABASE_URL и VITE_SUPABASE_ANON_KEY в .env.</div></div>}
      {error && <div className="error"><Info size={18}/><span>{error}</span><button onClick={refresh}><RefreshCw size={15}/> Повторить</button></div>}
      {message && <div className="success"><ShieldCheck size={17}/>{message}</div>}

      {loading ? <div className="empty">Загрузка…</div> : tab==="search" ? <>
        <div className="searchbox">
          <Search size={18}/>
          <input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Введи имя пловца…"/>
          {query && <button onClick={()=>setQuery("")}><X size={16}/></button>}
        </div>
        {!query.trim() ? <div className="empty"><Search size={30}/><p>{meets.length ? `Введи имя, чтобы найти результаты среди ${records.length} записей.` : "Пока нет загруженных протоколов."}</p></div>
        : query.trim().length<2 ? <div className="empty">Введи хотя бы 2 символа</div>
        : Object.keys(grouped).length===0 ? <div className="empty">Никого не нашли по запросу «{query}»</div>
        : Object.entries(grouped).slice(0,50).map(([name,results])=><SwimmerCard key={name} name={name} results={results} meetsById={meetsById}/>)}
      </> : <MeetBrowser meets={meets} onDelete={handleDelete}/>}

      <Rope/>

      <section className="upload">
        <UploadCloud size={24}/>
        <b>Добавить новый протокол</b>
        <small>Загрузи .xlsx или .xls файл с результатами — так, как присылает тренер</small>
        <label className="upload-button">Выбрать файл<input type="file" accept=".xlsx,.xls" onChange={handleFile}/></label>
        <div className="hint"><Info size={12}/> После загрузки результаты сохраняются в общей онлайн-базе.</div>
      </section>
    </main>
  </div>;
}