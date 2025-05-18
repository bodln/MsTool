using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace MsTool.Utlis
{
    public sealed class PibStore : IDisposable
    {
        private static readonly Lazy<PibStore> _lazy =
            new Lazy<PibStore>(() => new PibStore());

        public static PibStore Instance => _lazy.Value;

        private readonly SQLiteConnection _conn;
        private readonly SQLiteCommand _cmdUpsert;
        // In class cache for ease of access
        private readonly Dictionary<string, string> _cache;

        private PibStore()
        {
            var folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MsTool");

            Directory.CreateDirectory(folder);
            var dbFile = Path.Combine(folder, "pibs.sqlite");

            _conn = new SQLiteConnection($"Data Source={dbFile};");
            _conn.Open();

            using (var pragma = new SQLiteCommand("PRAGMA journal_mode = WAL;", _conn))
                pragma.ExecuteNonQuery();

            using (var cmd = new SQLiteCommand(@"
                CREATE TABLE IF NOT EXISTS Pibs (
                  PiB  TEXT PRIMARY KEY,
                  Name TEXT NOT NULL
                );", _conn))
            {
                cmd.ExecuteNonQuery();
            }

            _cmdUpsert = new SQLiteCommand(@"
                INSERT INTO Pibs (PiB, Name)
                  VALUES (@pib, @name)
                ON CONFLICT(PiB) DO UPDATE SET
                  Name = excluded.Name;", _conn);
            _cmdUpsert.Parameters.Add("@pib", System.Data.DbType.String);
            _cmdUpsert.Parameters.Add("@name", System.Data.DbType.String);

            _cache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = new SQLiteCommand("SELECT PiB,Name FROM Pibs;", _conn))
            using (var rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                    _cache[rdr.GetString(0)] = rdr.GetString(1);
            }
        }

        public string Lookup(string pib)
        {
            if (string.IsNullOrWhiteSpace(pib)) return null;
            _cache.TryGetValue(pib, out var name);
            return name;
        }

        public void AddOrUpdate(string pib, string name)
        {
            if (string.IsNullOrWhiteSpace(pib)) throw new ArgumentNullException(nameof(pib));
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));

            _cache[pib] = name;
            _cmdUpsert.Parameters["@pib"].Value = pib;
            _cmdUpsert.Parameters["@name"].Value = name;
            _cmdUpsert.ExecuteNonQuery();
        }

        public IReadOnlyDictionary<string, string> GetAll()
            => new Dictionary<string, string>(_cache);

        public void Dispose()
        {
            _cmdUpsert.Dispose();
            _conn.Dispose();
        }
    }
}
