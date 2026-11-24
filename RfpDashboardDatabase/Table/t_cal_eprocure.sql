CREATE TABLE public.t_cal_eprocure(
    id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    department_code VARCHAR(20) NOT NULL,
    department_name VARCHAR(255) NOT NULL,
    event_id VARCHAR(50) NOT NULL,
    event_name TEXT NOT NULL,
    format VARCHAR(50) NOT NULL,
    event_type VARCHAR(50) NOT NULL,
    end_date TIMESTAMPTZ NOT NULL,
    status VARCHAR(50) NOT NULL,
    buyer_name VARCHAR(255),
    buyer_email VARCHAR(255),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),    
    -- Optional: Enforce uniqueness on Department + Event ID
    CONSTRAINT uq_department_event UNIQUE (department_code, event_id)
);

ALTER TABLE public.t_cal_eprocure
    ADD COLUMN IF NOT EXISTS triage_status VARCHAR(30) NOT NULL DEFAULT 'PENDING',
    ADD COLUMN IF NOT EXISTS is_it BOOLEAN,
    ADD COLUMN IF NOT EXISTS llm_confidence NUMERIC(4,3),
    ADD COLUMN IF NOT EXISTS llm_reason TEXT,
    ADD COLUMN IF NOT EXISTS full_description TEXT,
    ADD COLUMN IF NOT EXISTS unspsc_codes TEXT,
    ADD COLUMN IF NOT EXISTS visited_at TIMESTAMPTZ;

CREATE INDEX IF NOT EXISTS idx_t_cal_eprocure_triage_status
    ON public.t_cal_eprocure (triage_status);