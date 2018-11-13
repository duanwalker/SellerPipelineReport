using System;
using System.IO;
using System.Collections.Generic;
using EllieMae.Encompass.Query;
using EllieMae.Encompass.Client;
using EllieMae.Encompass.Collections;
using EllieMae.Encompass.BusinessObjects.Loans;
using EllieMae.Encompass.BusinessObjects.Loans.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace Seller_Pipeline_Report
{
    class Program
    {
        static void Main(string[] args)
        {
            new EllieMae.Encompass.Runtime.RuntimeServices().Initialize();
            string SellerID = "";
            string LimitedSet = "";
            string newFileName = "";
            if (args.Length > 0)
            {
                SellerID = args[0];
                LimitedSet = args[1];
                newFileName = args[2];
            }

            ConditionReport(SellerID, LimitedSet, newFileName);
        }

        private static void ConditionReport(string SellerID, string LimitedSet, string newFileName)
        {
            try
            {
                //login to server
                Session session = new Session();
                session.Start("https://be11147937.ea.elliemae.net$be11147937", "sdkreport", "CqR3Fdt3LTKwVKCr");

                using (session)
                {
                    //build loan query
                    DateFieldCriterion tpoSubmitted = new DateFieldCriterion();
                    tpoSubmitted.FieldName = "Fields.TPO.X90";
                    tpoSubmitted.Value = DateFieldCriterion.NonEmptyDate;
                    tpoSubmitted.MatchType = OrdinalFieldMatchType.Equals;
                    tpoSubmitted.Precision = DateFieldMatchPrecision.Exact;

                    StringFieldCriterion sellerID = new StringFieldCriterion();
                    sellerID.FieldName = "Fields.TPO.X15";
                    sellerID.Value = SellerID;
                    sellerID.MatchType = StringFieldMatchType.Exact;
                    if (SellerID.ToUpper().Equals("ALL"))
                    {
                        sellerID.Value = "";
                        sellerID.Include = false;                        
                    }
                    else { sellerID.Include = true; }

                    NumericFieldCriterion allCondCount = new NumericFieldCriterion();
                    allCondCount.FieldName = "Fields.UWC.ALLCOUNT";
                    allCondCount.Value = 0;
                    allCondCount.MatchType = OrdinalFieldMatchType.GreaterThan;

                    StringFieldCriterion msPurchaseAC = new StringFieldCriterion();
                    msPurchaseAC.FieldName = "Fields.Log.MS.Status.Purchase - AC";
                    msPurchaseAC.Value = "Achieved";
                    msPurchaseAC.MatchType = StringFieldMatchType.Contains;
                    msPurchaseAC.Include = false;

                    StringFieldCriterion loanFolder = new StringFieldCriterion();
                    loanFolder.FieldName = "Loan.LoanFolder";
                    loanFolder.Value = "Corr. Active";
                    loanFolder.MatchType = StringFieldMatchType.Exact;
                    loanFolder.Include = true;

                    QueryCriterion query = loanFolder.And(sellerID.And(allCondCount.And(tpoSubmitted.And(msPurchaseAC))));

                    //get loans
                    LoanIdentityList ids = session.Loans.Query(query);
                    Console.WriteLine("Found " + ids.Count + " loans");
                    if (ids.Count > 0)
                    {
                        //Found Loans build Excel file
                        FileInfo newFile = new FileInfo(@"output\" + newFileName);
                        if (newFile.Exists)
                        {
                            newFile.Delete();
                            newFile = new FileInfo(@"output\"+newFileName);
                        }
                     
                        using (ExcelPackage package = new ExcelPackage(newFile))
                        {
                            //add a new worksheet
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Conditions");
                            //write header row
                            worksheet.Cells["A1"].Value = "Seller Name";
                            worksheet.Cells["B1"].Value = "Seller Loan Num";
                            worksheet.Cells["C1"].Value = "DH Loan Num";
                            worksheet.Cells["D1"].Value = "Borrower Name";
                            worksheet.Cells["E1"].Value = "Property State";
                            worksheet.Cells["F1"].Value = "Loan Purpose";
                            worksheet.Cells["G1"].Value = "Occupancy";
                            worksheet.Cells["H1"].Value = "Last Finished Milestone";
                            worksheet.Cells["I1"].Value = "Last Finished Milestone Date";
                            worksheet.Cells["J1"].Value = "Date Condition Added";
                            worksheet.Cells["K1"].Value = "Condition Status";
                            worksheet.Cells["L1"].Value = "Condition Type";
                            worksheet.Cells["M1"].Value = "Condition Title";
                            worksheet.Cells["N1"].Value = "Condition Details";
                            worksheet.Cells["O1"].Value = "Added By";
                            worksheet.Cells["P1"].Value = "For Intenal Use";
                            worksheet.Cells["Q1"].Value = "For External Use";
                            int row = 2;

                            foreach (LoanIdentity id in ids)
                            {
                                Loan loan = session.Loans.Open(id.Guid);
                                LogUnderwritingConditions conds = loan.Log.UnderwritingConditions;
                                //we only want the report to contain loans with conditions
                                if (conds.Count > 0)
                                {
                                    var ClosedList = new List<ConditionStatus> { (ConditionStatus)7, (ConditionStatus)8, (ConditionStatus)12 };
                                    Console.WriteLine("Found " + conds.Count + " total conditions in Loan: " + loan.LoanNumber);
                                    if (LimitedSet.ToUpper().Equals("LIMITED"))
                                    {
                                        Console.WriteLine("Only Pulling Open Conditions");
                                    }

                                    foreach (UnderwritingCondition cond in conds)
                                    {
                                        if (LimitedSet.ToUpper().Equals("ALL") || (!LimitedSet.ToUpper().Equals("ALL") && !ClosedList.Contains(cond.Status)))
                                        {
                                            //fill in row details here
                                            worksheet.Cells[row, 1].Value = loan.Fields["TPO.X14"];
                                            worksheet.Cells[row, 2].Value = loan.Fields["CX.DH.SELLERLOANNUM"];
                                            worksheet.Cells[row, 3].Value = loan.Fields["364"];
                                            worksheet.Cells[row, 4].Value = loan.Fields["4002"] + ", " + loan.Fields["4000"];
                                            worksheet.Cells[row, 5].Value = loan.Fields["14"];
                                            worksheet.Cells[row, 6].Value = loan.Fields["19"];
                                            worksheet.Cells[row, 7].Value = loan.Fields["1811"];
                                            worksheet.Cells[row, 8].Value = loan.Fields["LOG.MS.LASTCOMPLETED"];
                                            worksheet.Cells[row, 9].Value = loan.Fields["MS.STATUSDATE"].ToDate().Date;
                                            worksheet.Cells[row, 10].Value = cond.DateAdded.Date;
                                            worksheet.Cells[row, 11].Value = cond.Status.ToString();
                                            worksheet.Cells[row, 12].Value = cond.PriorTo.ToString();
                                            worksheet.Cells[row, 13].Value = cond.Title.ToString();
                                            worksheet.Cells[row, 14].Value = cond.Description.ToString().Replace("\"", "'");
                                            worksheet.Cells[row, 15].Value = cond.AddedBy.ToString();
                                            worksheet.Cells[row, 16].Value = cond.ForInternalUse.ToString();
                                            worksheet.Cells[row, 17].Value = cond.ForExternalUse.ToString();

                                            row++;
                                        }
                                    }
                                    row++;
                                }
                            }

                            //format range as table
                            var range = worksheet.Cells[1, 1, row - 2, 17];
                            var xltable = worksheet.Tables.Add(range, null);
                            xltable.TableStyle = TableStyles.Medium14;
                            //finish up worksheet
                            worksheet.Cells.AutoFitColumns(0);
                            worksheet.Column(9).Style.Numberformat.Format = "m/d/yyyy";
                            worksheet.Column(10).Style.Numberformat.Format = "m/d/yyyy";
                            worksheet.Column(14).Style.WrapText = true;
                            worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:1"];
                            worksheet.View.PageLayoutView = false;
                            worksheet.View.PageLayoutView = false;
                            //finish up package
                            package.Save();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unknown Error:");
                Console.WriteLine("\r\n\r\n{0}", ex);
            }

        } // End ConditionReport

    } //End Program
} //End Namespace