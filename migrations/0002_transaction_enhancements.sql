-- Transaction enhancements: preserve existing records while adding required fields and savings allocations.
-- Existing transfers remain readable but are not included in income/expense/savings totals.
ALTER TABLE transactions RENAME TO transactions_before_enhancements;

CREATE TABLE transactions (
  id TEXT PRIMARY KEY,
  user_id TEXT NOT NULL,
  account_id TEXT NOT NULL,
  category_id TEXT,
  transaction_type TEXT NOT NULL CHECK (transaction_type IN ('income', 'expense', 'savings_contribution', 'transfer')),
  amount_centavos INTEGER NOT NULL CHECK (amount_centavos > 0),
  transaction_date TEXT NOT NULL,
  description TEXT NOT NULL,
  transfer_group_id TEXT,
  payment_method TEXT NOT NULL DEFAULT 'cash' CHECK (payment_method IN ('cash', 'bank_transfer', 'e_wallet', 'credit_card', 'other')),
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

INSERT INTO transactions (id,user_id,account_id,category_id,transaction_type,amount_centavos,transaction_date,description,transfer_group_id,created_at,updated_at)
SELECT id,user_id,account_id,category_id,transaction_type,amount_centavos,transaction_date,COALESCE(description,''),transfer_group_id,created_at,updated_at
FROM transactions_before_enhancements;

DROP TABLE transactions_before_enhancements;
CREATE INDEX IF NOT EXISTS idx_transactions_user_date ON transactions(user_id, transaction_date);
CREATE INDEX IF NOT EXISTS idx_transactions_user_account ON transactions(user_id, account_id);
CREATE INDEX IF NOT EXISTS idx_transactions_user_type_date ON transactions(user_id, transaction_type, transaction_date);
