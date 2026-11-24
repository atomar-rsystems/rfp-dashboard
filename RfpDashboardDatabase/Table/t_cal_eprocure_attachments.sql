CREATE TABLE IF NOT EXISTS public.t_cal_eprocure_attachments (
    attachment_id SERIAL PRIMARY KEY,
    department_code VARCHAR(50) NOT NULL,
    event_id VARCHAR(50) NOT NULL,
    file_name TEXT NOT NULL,
    local_path TEXT NOT NULL,
    extracted_text TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT fk_event FOREIGN KEY (department_code, event_id)
        REFERENCES public.t_cal_eprocure (department_code, event_id) ON DELETE CASCADE
);