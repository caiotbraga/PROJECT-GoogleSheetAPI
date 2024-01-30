using CalculateApproval;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.IO;

string googleclientId = "1125466948302 34529805";
string googleclientSecret = "b8aa1fa7e1259f1ee0d4d19f80e716f0df54dfae\t";
string[] scopes = new[] { Google.Apis.Sheets.v4.SheetsService.Scope.Spreadsheets };
UserCredential credential = GoogleAuthentication.Login(googleclientId, googleclientSecret, scopes);

