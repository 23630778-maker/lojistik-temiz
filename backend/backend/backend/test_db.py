import psycopg

DB_URL = "postgres://postgres:Nisa2025Secure@db.mjnpmjfuinztssstvqsu.supabase.co:5432/postgres"

def test_connection():
    try:
        with psycopg.connect(DB_URL) as conn:
            with conn.cursor() as cur:
                cur.execute("INSERT INTO arac_kayitlari (plaka) VALUES (%s) RETURNING id;", ("TEST123",))
                inserted_id = cur.fetchone()[0]
        print(f"✅ Başarılı! Test kaydı eklendi, ID: {inserted_id}")
    except Exception as e:
        print(f"❌ HATA: {e}")

if __name__ == "__main__":
    test_connection()
