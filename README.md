This is application imports an excel file to Ms Sql database. First application reads the column names and validate them. If they contain unsupported characters, code will
replace them with acceptable chars. Then reads the rows to a dynamic object. DTO inserted to database via two repositories. One for entity framework and one for Dapper.
Repository pattern makes this application modular and expandable.
User can select excel files from the workin directory with an user interface. After importing, table from the database shown to screen as a table.
