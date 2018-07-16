use agenda 
-- alter table Users alter column senha varchar(40)
-- insert into users values(0,'Admaster2015','Admaster2015')

create table Users(
PK_User int primary key,
Login varchar(40) not null,
Senha varchar(20) not null
)

create table Contatos(
PK_Cont int primary key,
Nome varchar(80) not null,
Telefone varchar(20),
Celular varchar(20),
Genero varchar(20),
RedeSocial varchar(20),
FK_User int,
Foreign key(FK_User) references Users(PK_User))

create table Empresas(
PK_Emp int primary key,
Nome varchar(60) not null,
Login varchar(15),
Senha varchar(15)
)

create table Academia(
PK_Acm int primary key,
Nome varchar(20) not null,
Musculo varchar(20) not null
)

create table genero(
ID int primary key not null,
Descricao varchar(20) not null
)

create table social(
ID int primary key not null,
Descricao varchar(20) not null
)

insert into genero values(1,'Amigo')
insert into genero values(2,'Conhecido')
insert into genero values(3,'Familia')
insert into genero values(4,'Melhor amigo')

insert into social values(1,'Facebook')
insert into social values(2,'Vk')

select * from users
