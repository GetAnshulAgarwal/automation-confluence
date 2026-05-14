CREATE DATABASE config_api;
CREATE USER config_api_user WITH PASSWORD 'Ohpohshuquei1aechee1';
GRANT ALL PRIVILEGES ON DATABASE config_api TO config_api_user;
GRANT ALL ON SCHEMA public TO config_api_user;
GRANT ALL PRIVILEGES ON ALL TABLES IN SCHEMA public TO config_api_user;
ALTER DATABASE config_api OWNER TO config_api_user;
