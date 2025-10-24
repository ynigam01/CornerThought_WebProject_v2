-- Create the Users table
CREATE TABLE Users (
    UserID INTEGER PRIMARY KEY AUTOINCREMENT, -- AUTOINCREMENT for SQLite, use SERIAL for PostgreSQL
    UserName TEXT NOT NULL,
    UserOrganization TEXT NOT NULL,
    UserEmail TEXT UNIQUE NOT NULL
);

-- Create the UserCredentials table for secure password storage
CREATE TABLE UserCredentials (
    UserID INTEGER PRIMARY KEY,
    PasswordHash TEXT NOT NOT NULL,
    PasswordSalt TEXT NOT NULL,
    FOREIGN KEY (UserID) REFERENCES Users(UserID) ON DELETE CASCADE
);

-- Create the Organizations table
CREATE TABLE Organizations (
    OrganizationID INTEGER PRIMARY KEY AUTOINCREMENT, -- AUTOINCREMENT for SQLite, use SERIAL for PostgreSQL
    OrganizationName TEXT NOT NULL UNIQUE,
    OrganizationType TEXT,
    AdminName TEXT NOT NULL,
    AdminEmail TEXT NOT NULL,
    AllottedUsers INTEGER DEFAULT 0,
    DateAdded DATETIME DEFAULT CURRENT_TIMESTAMP,
    LastModified DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Optional: Create a table for password reset tokens
CREATE TABLE PasswordResetTokens (
    TokenID INTEGER PRIMARY KEY AUTOINCREMENT,
    UserID INTEGER NOT NULL,
    Token TEXT UNIQUE NOT NULL,
    ExpiresAt DATETIME NOT NULL,
    FOREIGN KEY (UserID) REFERENCES Users(UserID) ON DELETE CASCADE
);