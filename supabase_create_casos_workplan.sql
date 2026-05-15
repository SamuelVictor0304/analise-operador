create schema if not exists workplan;

drop table if exists workplan.casos_workplan cascade;

create table workplan.casos_workplan (
    id bigserial primary key,
    agreement_no varchar(50),
    cust_name varchar(255),
    dpd integer,
    total_amount_due numeric(15,2),
    last_contact_date date,
    allocation_date date,
    last_marking_date date,
    city varchar(150),
    state varchar(100),
    last_marking_value varchar(150),
    pct_of_margin_money numeric(10,4),
    no_first_ins_unpaid integer,
    cpf_cnpj varchar(30),
    disbursal_dealer_code varchar(50),
    dias_na_carteira integer,
    status varchar(20),
    regiao varchar(100),
    uf varchar(10),
    probabilidade varchar(100),
    supervisao varchar(100),
    faixa_atraso varchar(100),
    instalment_due_date date,
    flag_cobravel varchar(3),
    created_at timestamp without time zone default current_timestamp,
    updated_at timestamp without time zone default current_timestamp,
    status_base varchar(255),
    flag_cpc varchar(3),
    status_cpc varchar(30)
);
