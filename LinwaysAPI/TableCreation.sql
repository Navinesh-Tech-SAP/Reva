CREATE TABLE API_StudentFeeReceipt
(
    ID INT IDENTITY(1,1) PRIMARY KEY,

    AdmissionNo NVARCHAR(50),
    StudentName NVARCHAR(255),
    ReceiptNo NVARCHAR(100),
    ReceiptDate DATE,
    PaymentType NVARCHAR(50),

    FeeHeadName NVARCHAR(255),
    FeeHeadAmount DECIMAL(19,6),

    TotalAmount DECIMAL(19,6),

    -- SAP Integration Fields
    SAPDocEntry INT NULL,
    SAPDocNum INT NULL,

    IntegrationStatus NVARCHAR(20) DEFAULT 'Pending',
    IntegrationMessage NVARCHAR(MAX) NULL,

    CreatedDate DATETIME DEFAULT GETDATE(),
    UpdatedDate DATETIME NULL
);


--------------
to avaoid Dublicate
-----------------
CREATE UNIQUE INDEX IX_API_StudentFeeReceipt
ON API_StudentFeeReceipt(ReceiptNo, FeeHeadName);

--------------------
