const logActivity = (db, type, message) => {
  const stmt = db.prepare(
    'INSERT INTO activity (type, message) VALUES (@type, @message)'
  );
  stmt.run({ type, message });
};

const listActivity = (db, limit = 8) =>
  db
    .prepare(
      'SELECT id, type, message, created_at FROM activity ORDER BY created_at DESC LIMIT ?'
    )
    .all(limit);

module.exports = { logActivity, listActivity };
