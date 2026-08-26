/* ────────────────────────────────────────────────────────────────
   מחזור הסדנה - נקודת השינוי היחידה.

   הקובץ נטען גם ב-index.html וגם ב-thanks-webinar.html, כך
   שתאריך הוובינר נכתב פעם אחת ומופיע בשניהם.

   לפתיחת מחזור חדש מעדכנים כאן שלושה תאריכים. הכל בשעון
   ישראל; ההמרה ומעבר שעון קיץ/חורף מחושבים בזמן ריצה.
   ──────────────────────────────────────────────────────────────── */
window.CYCLE = {
  WEBINAR:          { date: '2026-10-04', time: '20:00', minutes: 60 },
  COURSE_START:     '2026-10-11',   // מפגש פתיחה
  LATE_JOIN_UNTIL:  '2026-10-18',   // המפגש השני - החלון האחרון למאחרים
  TIME:             '20:00',

  // קישור הזום לוובינר. כשהילה יוצרת את הפגישה - מדביקים אותו כאן,
  // וכל מייל תודה שנשלח מרגע הדחיפה יקבל אותו. ריק = המייל יכתוב
  // שהקישור יישלח לפני המפגש, במקום להבטיח קישור מת.
  ZOOM_LINK:        'https://us06web.zoom.us/j/86158650122?pwd=bfa9jDhZ3DRslUY3O2XtazELUrh60Y.1',

  NEXT_CYCLE_TEXT:  'בקרוב מועד נוסף',
  TZ:               'Asia/Jerusalem'
};

window.Cycle = (function () {
  const TZ = window.CYCLE.TZ;

  // היסט אזור הזמן של ישראל ברגע נתון, כולל שעון קיץ/חורף
  function tzOffset(ts) {
    const p = new Intl.DateTimeFormat('en-US', {
      timeZone: TZ, hour12: false, year: 'numeric', month: '2-digit',
      day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit'
    }).formatToParts(new Date(ts)).reduce((a, x) => (a[x.type] = x.value, a), {});
    return Date.UTC(+p.year, p.month - 1, +p.day,
                    p.hour % 24, +p.minute, +p.second) - ts;
  }

  // '2026-10-04' + '20:00' בשעון ישראל → רגע מוחלט
  function at(dateStr, timeStr) {
    const [Y, M, D] = String(dateStr).split('-').map(Number);
    const [hh, mm]  = String(timeStr || '00:00').split(':').map(Number);
    const naive = Date.UTC(Y, M - 1, D, hh, mm);
    let ts = naive;
    for (let i = 0; i < 2; i++) ts = naive - tzOffset(ts);
    return ts;
  }

  const fmtDate = ts => new Date(ts).toLocaleDateString('he-IL',
    { timeZone: TZ, weekday: 'long', day: 'numeric', month: 'numeric' });
  const short = ts => new Date(ts).toLocaleDateString('he-IL',
    { timeZone: TZ, day: 'numeric', month: 'numeric' });

  const C = window.CYCLE;
  return {
    at: at, fmtDate: fmtDate, short: short,
    webinarStart: at(C.WEBINAR.date, C.WEBINAR.time),
    webinarEnd:   at(C.WEBINAR.date, C.WEBINAR.time) + (C.WEBINAR.minutes || 60) * 60000,
    courseStart:  at(C.COURSE_START, C.TIME),
    lateEnd:      at(C.LATE_JOIN_UNTIL, C.TIME) + 60 * 60000
  };
})();
