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
