-- SQL script for creating the employees table
CREATE TABLE IF NOT EXISTS employees (
    name TEXT,
    age INTEGER,
    position TEXT,
    employee_id INTEGER,
    email TEXT,
    phone TEXT,
    address TEXT,
    department TEXT,
    salary REAL,
    start_date TEXT,
    supervisor TEXT,
    skills TEXT,
    education TEXT
);