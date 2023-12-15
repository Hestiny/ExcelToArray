using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using UnityEditor;
using UnityEngine;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Debug = UnityEngine.Debug;

namespace Norlo.ExcelToArray
{
    public class ExportExcelToArrayEditor : EditorWindow
    {
        [MenuItem("Tools/ExcelToArray/配置规则 ", priority = -200)]
        private static void OpenInfo()
        {
            var win = GetWindow<ExportExcelToArrayEditor>();
            win.name = "导出规则";
            win.minSize = new Vector2(400, 500);
            win.Show();
        }

        protected void OnGUI()
        {
            GUILayout.BeginVertical();
            GUILayout.TextArea("输出字典 每个表名为一个key 每个表为一个Array\n\n");
            GUILayout.EndVertical();
        }
    }

    public class ExcelToArray
    {
        private static readonly string _defaultExcelDirFullPath = Application.dataPath.Replace("Assets", "Excel");

        public struct SheetInfo
        {
            public ISheet Sheet;
            public List<List<string>> ValueList;

            public SheetInfo(ISheet sheet, List<List<string>> valueList)
            {
                Sheet = sheet;
                ValueList = valueList;
            }
        }

        private readonly Dictionary<string, SheetInfo> _excelInfoMap = new Dictionary<string, SheetInfo>();

        #region ====查询====

        /// <summary>
        /// 获取所有sheet名字
        /// </summary>
        /// <returns></returns>
        public List<string> GetSheetNames()
        {
            return _excelInfoMap.Keys.ToList();
        }

        public string GetCell(string sheet, int rowIndex, int colIndex)
        {
            if (!_excelInfoMap.TryGetValue(sheet, out var sheetInfo))
                return "";

            string cell;
            try
            {
                cell = sheetInfo.ValueList[rowIndex][colIndex];
            }
            catch (Exception)
            {
                Debug.LogErrorFormat("数据越界 sheetName={0}; row={1}; col={2}", sheet, rowIndex, colIndex);
                return "";
            }

            return cell;
        }

        public List<List<string>> GetSheet(string sheet)
        {
            List<List<string>> sheets = new List<List<string>>();
            return !_excelInfoMap.TryGetValue(sheet, out var sheetInfo) ? sheets : sheetInfo.ValueList;
        }

        /// <summary>
        /// 获取表内的所有合并单元
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public List<CellRangeAddress> GetMergeCell(string sheet)
        {
            List<CellRangeAddress> merges = new List<CellRangeAddress>();
            if (!_excelInfoMap.TryGetValue(sheet, out var sheetInfo))
                return merges;

            for (int i = 0; i < sheetInfo.Sheet.NumMergedRegions; i++)
            {
                merges.Add(sheetInfo.Sheet.GetMergedRegion(i));
            }

            return merges;
        }

        #endregion

        #region ====解析====

        /// <summary>
        /// 打开Excel文件夹
        /// </summary>
        public static void OpenExcelsFolder(string excelDirFullPath = null)
        {
            Process.Start("open", string.IsNullOrEmpty(excelDirFullPath) ? _defaultExcelDirFullPath : excelDirFullPath);
        }

        /// <summary>
        /// 导出Excel到cs
        /// </summary>
        /// <param name="excelFullPath">Excel存放的文件目录</param>
        public static ExcelToArray ExportExcels(string excelFullPath = null)
        {
            ExcelToArray excelToCs = new ExcelToArray();
            excelToCs._excelInfoMap.Clear();
            string[] tablePaths = Directory.GetFiles(string.IsNullOrEmpty(excelFullPath) ? _defaultExcelDirFullPath : excelFullPath, "*",
                SearchOption.TopDirectoryOnly);
            foreach (var fileFullPath in tablePaths)
            {
                excelToCs.ParseExcel(fileFullPath);
            }

            return excelToCs;
        }

        #endregion

        private void ParseExcel(string fileFullPath)
        {
            using FileStream stream = new FileStream(fileFullPath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = null;
            if (fileFullPath.EndsWith(".xlsx"))
                workbook = new XSSFWorkbook(fileFullPath); //2007
            else if (fileFullPath.EndsWith(".xls"))
                workbook = new HSSFWorkbook(stream); //2003

            if (workbook == null) return;
            int sheetNumber = workbook.NumberOfSheets;
            for (int sheetIndex = 0; sheetIndex < sheetNumber; sheetIndex++)
            {
                string sheetName = workbook.GetSheetName(sheetIndex);
                List<List<string>> pickInfoList = new List<List<string>>();
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                sheet.ForceFormulaRecalculation = true; //强制公式计算

                int maxColumnNum = sheet.GetRow(0).LastCellNum;
                for (int rowId = 0; rowId <= sheet.LastRowNum; rowId++)
                {
                    List<string> rowInfoList = new List<string>();
                    IRow sheetRowInfo = sheet.GetRow(rowId);
                    if (sheetRowInfo == null)
                    {
                        Debug.LogErrorFormat("无法获取行数据 sheetName={0} ;rowId={1};rowMax={2}", sheetName, rowId, sheet.LastRowNum);
                    }

                    if (sheetRowInfo != null)
                    {
                        var rowFirstCell = sheetRowInfo.GetCell(0);
                        if (null == rowFirstCell)
                            continue;
                        if (rowFirstCell.CellType == CellType.Blank || rowFirstCell.CellType == CellType.Unknown ||
                            rowFirstCell.CellType == CellType.Error)
                            continue;
                    }

                    for (int columnId = 0; columnId < maxColumnNum; columnId++)
                    {
                        if (sheetRowInfo == null) continue;
                        ICell pickCell = sheetRowInfo.GetCell(columnId);

                        switch (pickCell)
                        {
                            case { IsMergedCell: true }:
                                pickCell = GetMergeCell(sheet, pickCell.RowIndex, pickCell.ColumnIndex);
                                break;
                            case null:
                                pickCell = GetMergeCell(sheet, rowId, columnId);
                                break;
                        }

                        if (pickCell is { CellType: CellType.Formula })
                        {
                            pickCell.SetCellType(CellType.String);
                            rowInfoList.Add(pickCell.StringCellValue);
                        }
                        else if (pickCell != null)
                        {
                            rowInfoList.Add(pickCell.ToString());
                        }
                        else
                        {
                            rowInfoList.Add("");
                        }
                    }

                    pickInfoList.Add(rowInfoList);
                }

                if (!_excelInfoMap.TryAdd(sheetName, new SheetInfo(sheet, pickInfoList)))
                {
                    Debug.LogErrorFormat("sheetName:{0} 重复了!", sheetName);
                }
            }
        }

        /// <summary>
        /// 获取合并的格子的原格子
        /// 合并格子的首行首列就是合并单元格的信息
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <returns></returns>
        private static ICell GetMergeCell(ISheet sheet, int rowIndex, int colIndex)
        {
            for (int ii = 0; ii < sheet.NumMergedRegions; ii++)
            {
                var cellRange = sheet.GetMergedRegion(ii);
                if (colIndex < cellRange.FirstColumn ||
                    colIndex > cellRange.LastColumn ||
                    rowIndex < cellRange.FirstRow ||
                    rowIndex > cellRange.LastRow)
                    continue;
                var row = sheet.GetRow(cellRange.FirstRow);
                var mergeCell = row.GetCell(cellRange.FirstColumn);

                return mergeCell;
            }

            return null;
        }
    }
}