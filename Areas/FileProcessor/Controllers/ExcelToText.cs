using ExcelProcessor;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using NuGet.Versioning;
using System.IO;
using System.Text;

namespace ExcelProcessor.Areas.FileProcessor.Controllers
{
    [Area("FileProcessor")]
    public class ExcelToText : Controller
    {
        public IActionResult Index()
        {
            return View();
        }


        [HttpPost]
        public IActionResult UploadAndConvert(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ViewBag.ErrorMessage = "Please select a file for upload.";
                return View("Index");
            }

            try
            {




                using (var stream = file.OpenReadStream())
                {
                    stream.Position = 0;  // Ensure stream is at the beginning
                    using (var workbook = new XSSFWorkbook(stream))
                    {
                        var worksheet = workbook.GetSheetAt(0); // Assuming the data is in the first sheet
                        var rowCount = worksheet.PhysicalNumberOfRows;

                        var stringBuilder = new StringBuilder();

                        var headerRow = worksheet.GetRow(0);
                        ExcelColumnIndex excelColumnIndex = InitExcelColIndex(headerRow);

                        // Iterate through rows and columns
                        for (int row = 1; row < rowCount; row++)
                        {
                            var currentRow = worksheet.GetRow(row);
                            if (currentRow != null)
                            {




                                // Get the cell value
                                if (excelColumnIndex.CustomerID > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.CustomerID)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Applicant_Type > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Applicant_Type)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Prefix > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Prefix)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.First_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.First_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Middle_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Middle_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Last_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Last_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Father_Name_Prefix > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Father_Name_Prefix)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Father_First_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Father_First_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Father_Middle_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Father_Middle_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Father_Last_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Father_Last_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Mother_Prefix > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Mother_Prefix)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Mother_First_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Mother_First_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Mother_Middle_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Mother_Middle_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Mother_Last_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Mother_Last_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Date_Of_Birth > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Date_Of_Birth)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Gender > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Gender)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Marital_Status > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Marital_Status)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Citizenship > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Citizenship)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Residential_status > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Residential_status)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Occupation_Type > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Occupation_Type)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Passport_No > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Passport_No)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Passport_ExpireDate > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Passport_ExpireDate)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Voter_ID_Card > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Voter_ID_Card)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.PAN_Card > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.PAN_Card)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Driving_Licence > 0) { stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Driving_Licence)?.ToString() ?? string.Empty); stringBuilder.Append("|"); }
                                if (excelColumnIndex.Driving_Licence_Expiry_Date > 0) { stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Driving_Licence_Expiry_Date)?.ToString() ?? string.Empty); stringBuilder.Append("|"); }
                                if (excelColumnIndex.Adhar_Card_NO > 0) { stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Adhar_Card_NO)?.ToString() ?? string.Empty); stringBuilder.Append("|"); }
                                if (excelColumnIndex.Addresstype > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Addresstype)?.ToString() ?? string.Empty);
                                    stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.CP_Line2 > 0)
                                {

                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.CP_Line2)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.CP_Line3 > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.CP_Line3)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.CP_City > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.CP_City)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.CP_District > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.CP_District)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.CP_Pincode > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.CP_Pincode)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.CP_State_Code > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.CP_State_Code)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.CP_ISO_3166_Country_Code > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.CP_ISO_3166_Country_Code)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.STDCode > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.STDCode)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Mobile_No > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Mobile_No)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Email_ID > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Email_ID)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Applicant_Declaration_Date > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Applicant_Declaration_Date)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.Applicant_Declaration_Place > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.Applicant_Declaration_Place)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.KYC_Verification_Date > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.KYC_Verification_Date)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.kyc_verification_name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.kyc_verification_name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.KYC_Verification_Emp_Code > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.KYC_Verification_Emp_Code)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.KYC_Verification_Emp_Designation > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.KYC_Verification_Emp_Designation)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.KYC_Verification_Emp_Branch > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.KYC_Verification_Emp_Branch)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.INSTITUTION_Name > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.INSTITUTION_Name)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                }
                                if (excelColumnIndex.INSTITUTION_Code > 0)
                                {
                                    stringBuilder.Append(currentRow.GetCell(excelColumnIndex.INSTITUTION_Code)?.ToString() ?? string.Empty); stringBuilder.Append("|");
                                } 

                                stringBuilder.AppendLine();
                            }
                        }

                        // Set the content type and file name for the response
                        var fileName = "output.txt";
                        var contentType = "text/plain";

                        // Convert the text content to bytes
                        var fileBytes = Encoding.UTF8.GetBytes(stringBuilder.ToString());

                        // Return the text content as a downloadable file
                        return File(fileBytes, contentType, fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = $"An error occurred: {ex.Message}";
                return View("Index");
            }
        }

        private ExcelColumnIndex InitExcelColIndex(IRow headerRow)
        {
            ExcelColumnIndex excelColumnIndex = new ExcelColumnIndex();
            excelColumnIndex.CustomerID = GetCellIndexByColumnName(headerRow, ExcelFileConfig.CustomerID);

            excelColumnIndex.Applicant_Type = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Applicant_Type);

            excelColumnIndex.Prefix = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Prefix);

            excelColumnIndex.First_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.First_Name);

            excelColumnIndex.Middle_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Middle_Name);

            excelColumnIndex.Last_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Last_Name);

            excelColumnIndex.Father_Name_Prefix = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Father_Name_Prefix);

            excelColumnIndex.Father_First_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Father_First_Name);

            excelColumnIndex.Father_Middle_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Mother_Middle_Name);

            excelColumnIndex.Father_Last_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Mother_Prefix);

            excelColumnIndex.Mother_Prefix = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Mother_First_Name);

            excelColumnIndex.Mother_First_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Mother_Middle_Name);

            excelColumnIndex.Mother_Middle_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Mother_Last_Name);

            excelColumnIndex.Mother_Last_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Mother_Last_Name);

            excelColumnIndex.Date_Of_Birth = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Date_Of_Birth);

            excelColumnIndex.Gender = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Gender);

            excelColumnIndex.Marital_Status = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Marital_Status);

            excelColumnIndex.Citizenship = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Citizenship);

            excelColumnIndex.Residential_status = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Residential_status);

            excelColumnIndex.Occupation_Type = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Occupation_Type);

            excelColumnIndex.Passport_No = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Passport_No);

            excelColumnIndex.Passport_ExpireDate = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Passport_ExpireDate);

            excelColumnIndex.Voter_ID_Card = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Voter_ID_Card);

            excelColumnIndex.PAN_Card = GetCellIndexByColumnName(headerRow, ExcelFileConfig.PAN_Card);

            excelColumnIndex.Driving_Licence = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Driving_Licence);

            excelColumnIndex.Driving_Licence_Expiry_Date = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Driving_Licence_Expiry_Date);

            excelColumnIndex.Adhar_Card_NO = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Adhar_Card_NO);

            excelColumnIndex.Addresstype = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Addresstype);

            excelColumnIndex.CP_Line2 = GetCellIndexByColumnName(headerRow, ExcelFileConfig.CP_Line2);

            excelColumnIndex.CP_Line3 = GetCellIndexByColumnName(headerRow, ExcelFileConfig.CP_Line3);

            excelColumnIndex.CP_City = GetCellIndexByColumnName(headerRow, ExcelFileConfig.CP_City);

            excelColumnIndex.CP_District = GetCellIndexByColumnName(headerRow, ExcelFileConfig.CP_District);

            excelColumnIndex.CP_Pincode = GetCellIndexByColumnName(headerRow, ExcelFileConfig.CP_Pincode);

            excelColumnIndex.CP_State_Code = GetCellIndexByColumnName(headerRow, ExcelFileConfig.CP_State_Code);

            excelColumnIndex.CP_ISO_3166_Country_Code = GetCellIndexByColumnName(headerRow, ExcelFileConfig.CP_ISO_3166_Country_Code);

            excelColumnIndex.STDCode = GetCellIndexByColumnName(headerRow, ExcelFileConfig.STDCode);

            excelColumnIndex.Mobile_No = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Mobile_No);

            excelColumnIndex.Email_ID = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Email_ID);

            excelColumnIndex.Applicant_Declaration_Date = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Applicant_Declaration_Date);

            excelColumnIndex.Applicant_Declaration_Place = GetCellIndexByColumnName(headerRow, ExcelFileConfig.Applicant_Declaration_Place);

            excelColumnIndex.KYC_Verification_Date = GetCellIndexByColumnName(headerRow, ExcelFileConfig.KYC_Verification_Date);

            excelColumnIndex.kyc_verification_name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.kyc_verification_name);

            excelColumnIndex.KYC_Verification_Emp_Code = GetCellIndexByColumnName(headerRow, ExcelFileConfig.KYC_Verification_Emp_Code);

            excelColumnIndex.KYC_Verification_Emp_Designation = GetCellIndexByColumnName(headerRow, ExcelFileConfig.KYC_Verification_Emp_Designation);

            excelColumnIndex.KYC_Verification_Emp_Branch = GetCellIndexByColumnName(headerRow, ExcelFileConfig.KYC_Verification_Emp_Branch);

            excelColumnIndex.INSTITUTION_Name = GetCellIndexByColumnName(headerRow, ExcelFileConfig.INSTITUTION_Name);

            excelColumnIndex.INSTITUTION_Code = GetCellIndexByColumnName(headerRow, ExcelFileConfig.INSTITUTION_Code);
            return excelColumnIndex;
        }
        public static int GetCellIndexByColumnName(IRow row, string columnName)
        {
            int columnIndex = -1;

            for (int i = 0; i < row.LastCellNum; i++)
            {
                NPOI.SS.UserModel.ICell cell = row.GetCell(i);
                if (cell != null && cell.StringCellValue == columnName)
                {
                    columnIndex = i;
                    break;
                }
            }

            return columnIndex;
        }
    }
}
