-- Budget categories, goal descriptions, and linked savings contributions.
-- Additive migration: preserves existing records and defaults legacy categories to Needs.
ALTER TABLE categories ADD COLUMN budget_group TEXT NOT NULL DEFAULT 'needs' CHECK (budget_group IN ('needs','wants','savings','income'));
ALTER TABLE savings_goals ADD COLUMN description TEXT;
ALTER TABLE transactions ADD COLUMN goal_id TEXT;

CREATE INDEX IF NOT EXISTS idx_transactions_user_goal_date ON transactions(user_id, goal_id, transaction_date);
CREATE INDEX IF NOT EXISTS idx_categories_user_group ON categories(user_id, budget_group);
