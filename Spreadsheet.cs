// <copyright file="Spreadsheet.cs" company="UofU-CS3500">
// Copyright (c) 2024 UofU-CS3500. All rights reserved.
// </copyright>
// <authors>LAN QUANG HUYNH</authors>
// <date>10/04/2024</date>/

// Written by Joe Zachary for CS 3500, September 2013
// Update by Profs Kopta and de St. Germain and LAN QUANG HUYNH
namespace CS3500.Spreadsheet;

using CS3500.Formula;
using CS3500.DependencyGraph;
using System.Text.RegularExpressions;
using System.Text.Json;

/// <summary>
/// Author:    LAN QUANG HUYNH
/// Partner:   None
/// Date:      09/23/2024
/// Course:    CS 3500, University of Utah, School of Computing
/// Copyright: CS 3500 and LAN QUANG HUYNH - This work may not
///            be copied for use in Academic Coursework.
///
/// I, LAN QUANG HUYNH, certify that I wrote this code from scratch and
/// did not copy it in part or whole from another source.  All
/// references used in the completion of the assignments are cited
/// in my README file.
///
/// This file contains the implementation of the Spreadsheet class, which represents the core data structure
/// of the spreadsheet application. It handles cell storage, formula evaluation, and dependencies between cells.
///
/// Key Responsibilities:
///     - Storing cell values, including numbers, text, and formulas.
///     - Handling the evaluation of cell formulas, taking into account dependencies between cells.
///     - Managing the recalculation of dependent cells when a cell value is updated.
///     - Supporting various operations such as saving, loading, and validation of spreadsheet content.
///
/// This class is designed to efficiently handle updates and ensure that cells are correctly evaluated
/// in dependency order. Exceptions are thrown in cases of circular dependencies or invalid formulas.
/// </summary>

/// <summary>
///   <para>
///     Thrown to indicate that a change to a cell will cause a circular dependency.
///   </para>
/// </summary>
public class CircularException : Exception
{
}

/// <summary>
///   <para>
///     Thrown to indicate that a name parameter was invalid.
///   </para>
/// </summary>
public class InvalidNameException : Exception
{
}

/// <summary>
/// <para>
///   Thrown to indicate that a read or write attempt has failed with
///   an expected error message informing the user of what went wrong.
/// </para>
/// </summary>
public class SpreadsheetReadWriteException : Exception
{
    /// <summary>
    /// Initializes a new instance of the <see cref="SpreadsheetReadWriteException"/> class.
    ///   <para>
    ///     Creates the exception with a message defining what went wrong.
    ///   </para>
    /// </summary>
    /// <param name="msg"> An informative message to the user. </param>
    public SpreadsheetReadWriteException(string msg)
    : base(msg)
    {
    }
}

/// <summary>
///   <para>
///     An Spreadsheet object represents the state of a simple spreadsheet.  A
///     spreadsheet represents an infinite number of named cells.
///   </para>
/// <para>
///     Valid Cell Names: A string is a valid cell name if and only if it is one or
///     more letters followed by one or more numbers, e.g., A5, BC27.
/// </para>
/// <para>
///    Cell names are case insensitive, so "x1" and "X1" are the same cell name.
///    Your code should normalize (uppercased) any stored name but accept either.
/// </para>
/// <para>
///     A spreadsheet represents a cell corresponding to every possible cell name.  (This
///     means that a spreadsheet contains an infinite number of cells.)  In addition to
///     a name, each cell has a contents and a value.  The distinction is important.
/// </para>
/// <para>
///     The <b>contents</b> of a cell can be (1) a string, (2) a double, or (3) a Formula.
///     If the contents of a cell is set to the empty string, the cell is considered empty.
/// </para>
/// <para>
///     By analogy, the contents of a cell in Excel is what is displayed on
///     the editing line when the cell is selected.
/// </para>
/// <para>
///     In a new spreadsheet, the contents of every cell is the empty string. Note:
///     this is by definition (it is IMPLIED, not stored).
/// </para>
/// <para>
///     The <b>value</b> of a cell can be (1) a string, (2) a double, or (3) a FormulaError.
///     (By analogy, the value of an Excel cell is what is displayed in that cell's position
///     in the grid.)
/// </para>
/// <list type="number">
///   <item>If a cell's contents is a string, its value is that string.</item>
///   <item>If a cell's contents is a double, its value is that double.</item>
///   <item>
///     <para>
///       If a cell's contents is a Formula, its value is either a double or a FormulaError,
///       as reported by the Evaluate method of the Formula class.  For this assignment,
///       you are not dealing with values yet.
///     </para>
///   </item>
/// </list>
/// <para>
///     Spreadsheets are never allowed to contain a combination of Formulas that establish
///     a circular dependency.  A circular dependency exists when a cell depends on itself.
///     For example, suppose that A1 contains B1*2, B1 contains C1*2, and C1 contains A1*2.
///     A1 depends on B1, which depends on C1, which depends on A1.  That's a circular
///     dependency.
/// </para>
/// </summary>
public class Spreadsheet
{
    /// <summary>
    /// A dictionary that maps cell names (e.g., "A1", "B2") to their corresponding Cell objects.
    /// </summary>
    private readonly Dictionary<string, Cell> cells;

    /// <summary>
    /// An instance of the DependencyGraph class, which manages the relationships between cells.
    /// </summary>
    private readonly DependencyGraph dg;

    /// <summary>
    /// A name of the spreadshheet.
    /// </summary>
    // A name of the spreadshheet
    private readonly string name;

    /// <summary>
    /// Initializes a new instance of the <see cref="Spreadsheet"/> class.
    /// </summary>
    public Spreadsheet()
    {
        this.cells = new Dictionary<string, Cell>();
        this.dg = new DependencyGraph();
        this.name = "default";
        this.Changed = false;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Spreadsheet"/> class with the specified name.
    /// </summary>
    /// <param name="name">The name of the spreadsheet to be initialized.</param>
    public Spreadsheet(string name)
    {
        this.cells = new Dictionary<string, Cell>();
        this.dg = new DependencyGraph();
        this.name = name;
        this.Changed = false;
    }

    /// <summary>
    /// Gets a value indicating whether the spreadsheet was modified
    /// True if this spreadsheet has been modified since it was created or saved
    /// (whichever happened most recently); false otherwise.
    /// </summary>
    public bool Changed { get; private set; }

    /// <summary>
    ///   <para>
    ///     Shortcut syntax to for getting the value of the cell
    ///     using the [] operator.
    ///   </para>
    ///   <para>
    ///     See: <see cref="GetCellValue(string)"/>.
    ///   </para>
    ///   <para>
    ///     Example Usage:
    ///   </para>
    ///   <code>
    ///      sheet.SetContentsOfCell( "A1", "=5+5" );
    ///
    ///      sheet["A1"] == 10;
    ///      // vs.
    ///      sheet.GetCellValue("A1") == 10;
    ///   </code>
    /// </summary>
    /// <param name="cellName"> Any valid cell name. </param>
    /// <returns>
    ///   Returns the value of a cell.  Note: If the cell is a formula, the value should
    ///   already have been computed.
    /// </returns>
    /// <exception cref="InvalidNameException">
    ///     If the name parameter is invalid, throw an InvalidNameException.
    /// </exception>
    public object this[string cellName]
    {
        get
        {
            // Call GetCellValue to retrieve the value
            return GetCellValue(cellName);
        }
    }

    /// <summary>
    ///   Provides a copy of the names of all of the cells in the spreadsheet
    ///   that contain information (i.e., not empty cells).
    /// </summary>
    /// <returns>
    ///   A set of the names of all the non-empty cells in the spreadsheet.
    /// </returns>
    public ISet<string> GetNamesOfAllNonemptyCells()
    {
        return new HashSet<string>(this.cells.Keys);
    }

    /// <summary>
    ///   Returns the contents (as opposed to the value) of the named cell.
    /// </summary>
    ///
    /// <exception cref="InvalidNameException">
    ///   Thrown if the name is invalid.
    /// </exception>
    ///
    /// <param name="name">The name of the spreadsheet cell to query. </param>
    /// <returns>
    ///   The contents as either a string, a double, or a Formula.
    ///   See the class header summary.
    /// </returns>
    public object GetCellContents(string name)
    {
        name = name.ToUpper();
        this.ValidateCellName(name);

        Cell value;
        if (this.cells.ContainsKey(name))
        {
            value = this.cells[name];
            if (value.Contents is Formula formula)
            {
                return formula;
            }
            else if (value.Contents is double number)
            {
                return number;
            }
            else if(value.Contents is string stringContent)
            {
                return stringContent;
            }
        }

        return string.Empty;
    }

    /// <summary>
    ///   <para>
    ///     Return the value of the named cell.
    ///   </para>
    /// </summary>
    /// <param name="cellName"> The cell in question. </param>
    /// <returns>
    ///   Returns the value (as opposed to the contents) of the named cell.  The return
    ///   value's type should be either a string, a double, or a CS3500.Formula.FormulaError.
    ///   If the cell contents are a formula, the value should have already been computed
    ///   at this point.
    /// </returns>
    /// <exception cref="InvalidNameException">
    ///   If the provided name is invalid, throws an InvalidNameException.
    /// </exception>
    public object GetCellValue(string cellName)
    {
        cellName = cellName.ToUpper();

        // if name is invalid, throw exception
        this.ValidateCellName(cellName);

        // Otherwise return the value of the named cell
        if (cells.ContainsKey(cellName))
        {
            return cells[cellName].Value;
        }

        return string.Empty;
    }

    /// <summary>
    /// Writes the contents of this spreadsheet in JSON form.
    /// </summary>
    /// <returns>The string in JSON form that contains the contents of the spreadsheet.</returns>
    public string GetJSON()
    {
        // Create a dictionary for storing cells and their string representation
        Dictionary<string, Dictionary<string, string>> cellsDict = new Dictionary<string, Dictionary<string, string>>();

        foreach (var cell in cells)
        {
            string cellContent = string.Empty;
            if (cell.Value.Contents is Formula formula)
            {
                cellContent = "=" + formula.ToString(); // Formulas are serialized with an '=' prefix
            }
            else if (cell.Value.Contents is double number)
            {
                cellContent = number.ToString(); // Doubles are serialized as is
            }
            else if (cell.Value.Contents is string stringContent)
            {
                cellContent = stringContent; // Strings are serialized as is
            }

            // Add to dictionary with the expected "StringForm"
            cellsDict[cell.Key] = new Dictionary<string, string>
                {
                    { "StringForm", cellContent },
                };
        }

        var jsonObject = new { Cells = cellsDict };

        // Serialize to JSON
        string json = JsonSerializer.Serialize(jsonObject, new JsonSerializerOptions { WriteIndented = true });

        // Mark the spreadsheet as no longer "changed"
        this.Changed = false;

        return json;
    }

    /// <summary>
    ///   <para>
    ///     Writes the contents of this spreadsheet to the named file using a JSON format.
    ///     If the file already exists, overwrite it.
    ///   </para>
    ///   <para>
    ///     The output JSON should look like the following.
    ///   </para>
    ///   <para>
    ///     For example, consider a spreadsheet that contains a cell "A1"
    ///     with contents being the double 5.0, and a cell "B3" with contents
    ///     being the Formula("A1+2"), and a cell "C4" with the contents "hello".
    ///   </para>
    ///   <para>
    ///      This method would produce the following JSON string:
    ///   </para>
    ///   <code>
    ///   {
    ///     "Cells": {
    ///       "A1": {
    ///         "StringForm": "5"
    ///       },
    ///       "B3": {
    ///         "StringForm": "=A1+2"
    ///       },
    ///       "C4": {
    ///         "StringForm": "hello"
    ///       }
    ///     }
    ///   }
    ///   </code>
    ///   <para>
    ///     You can achieve this by making sure your data structure is a dictionary
    ///     and that the contained objects (Cells) have property named "StringForm"
    ///     (if this name does not match your existing code, use the JsonPropertyName
    ///     attribute).
    ///   </para>
    ///   <para>
    ///     There can be 0 cells in the dictionary, resulting in { "Cells" : {} }.
    ///   </para>
    ///   <para>
    ///     Further, when writing the value of each cell...
    ///   </para>
    ///   <list type="bullet">
    ///     <item>
    ///       If the contents is a string, the value of StringForm is that string
    ///     </item>
    ///     <item>
    ///       If the contents is a double d, the value of StringForm is d.ToString()
    ///     </item>
    ///     <item>
    ///       If the contents is a Formula f, the value of StringForm is "=" + f.ToString()
    ///     </item>
    ///   </list>
    ///   <para>
    ///     After saving the file, the spreadsheet is no longer "changed".
    ///   </para>
    /// </summary>
    /// <param name="filename"> The name (with path) of the file to save to.</param>
    /// <exception cref="SpreadsheetReadWriteException">
    ///   If there are any problems opening, writing, or closing the file,
    ///   the method should throw a SpreadsheetReadWriteException with an
    ///   explanatory message.
    /// </exception>
    public void Save(string filename)
    {
        try
        {
            // Serialize to JSON
            string json = this.GetJSON();

            // Write to file
            File.WriteAllText(filename, json);

            // Mark the spreadsheet as no longer "changed"
            this.Changed = false;
        }
        catch (Exception ex)
        {
            throw new SpreadsheetReadWriteException("Error saving the spreadsheet: " + ex.Message);
        }
    }

    /// <summary>
    ///   <para>
    ///     Read the data (JSON) from the file and instantiate the current
    ///     spreadsheet.  See <see cref="Save(string)"/> for expected format.
    ///   </para>
    ///   <para>
    ///     Note: First deletes any current data in the spreadsheet.
    ///   </para>
    ///   <para>
    ///     Loading a spreadsheet should set changed to false.  External
    ///     programs should alert the user before loading over a changed sheet.
    ///   </para>
    /// </summary>
    /// <param name="filename"> The saved file name including the path. </param>
    /// <exception cref="SpreadsheetReadWriteException"> When the file cannot be opened or the json is bad.</exception>
    public void Load(string filename)
    {
        try
        {
            // Read the file and deserialize the JSON into the appropriate structure
            string json = File.ReadAllText(filename);
            var jsonObject = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, Dictionary<string, string>>>>(json);

            // Clear current data
            this.cells.Clear();
            if (jsonObject != null)
            {
                // Load the new cells
                foreach (KeyValuePair<string, Dictionary<string, string>> cell in jsonObject["Cells"])
                {
                    string cellName = cell.Key;
                    string content = cell.Value["StringForm"];

                    // Determine the type of the content and set it accordingly
                    if (content.StartsWith("="))
                    {
                        // Formula case
                        this.SetContentsOfCell(cellName, content);
                    }
                    else if (double.TryParse(content, out double number))
                    {
                        // Double case
                        this.SetContentsOfCell(cellName, number.ToString());
                    }
                    else
                    {
                        // String case
                        this.SetContentsOfCell(cellName, content);
                    }
                }
            }

            // Mark the spreadsheet as no longer "changed"
            this.Changed = false;
        }
        catch (Exception ex)
        {
            throw new SpreadsheetReadWriteException("Error loading the spreadsheet: " + ex.Message);
        }
    }

    /// <summary>
    ///   <para>
    ///       Sets the contents of the named cell to the appropriate object
    ///       based on the string in <paramref name="content"/>.
    ///   </para>
    ///   <para>
    ///       First, if the <paramref name="content"/> parses as a double, the contents of the named
    ///       cell becomes that double.
    ///   </para>
    ///   <para>
    ///       Otherwise, if the <paramref name="content"/> begins with the character '=', an attempt is made
    ///       to parse the remainder of content into a Formula.
    ///   </para>
    ///   <para>
    ///       There are then three possible outcomes when a formula is detected:
    ///   </para>
    ///
    ///   <list type="number">
    ///     <item>
    ///       If the remainder of content cannot be parsed into a Formula, a
    ///       FormulaFormatException is thrown.
    ///     </item>
    ///     <item>
    ///       If changing the contents of the named cell to be f
    ///       would cause a circular dependency, a CircularException is thrown,
    ///       and no change is made to the spreadsheet.
    ///     </item>
    ///     <item>
    ///       Otherwise, the contents of the named cell becomes f.
    ///     </item>
    ///   </list>
    ///   <para>
    ///     Finally, if the content is a string that is not a double and does not
    ///     begin with an "=" (equal sign), save the content as a string.
    ///   </para>
    ///   <para>
    ///     On successfully changing the contents of a cell, the spreadsheet will be <see cref="Changed"/>.
    ///   </para>
    /// </summary>
    /// <param name="name"> The cell name that is being changed.</param>
    /// <param name="content"> The new content of the cell.</param>
    /// <returns>
    ///   <para>
    ///     This method returns a list consisting of the passed in cell name,
    ///     followed by the names of all other cells whose value depends, directly
    ///     or indirectly, on the named cell. The order of the list MUST BE any
    ///     order such that if cells are re-evaluated in that order, their dependencies
    ///     are satisfied by the time they are evaluated.
    ///   </para>
    ///   <para>
    ///     For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
    ///     list {A1, B1, C1} is returned.  If the cells are then evaluate din the order:
    ///     A1, then B1, then C1, the integrity of the Spreadsheet is maintained.
    ///   </para>
    /// </returns>
    /// <exception cref="InvalidNameException">
    ///   If the name parameter is invalid, throw an InvalidNameException.
    /// </exception>
    /// <exception cref="CircularException">
    ///   If changing the contents of the named cell to be the formula would
    ///   cause a circular dependency, throw a CircularException.
    ///   (NOTE: No change is made to the spreadsheet.)
    /// </exception>
    public IList<string> SetContentsOfCell(string name, string content)
    {
        IList<string> all_dependents;
        object parsedContent = ParseContent(content);

        if (parsedContent is double number)
        {
            all_dependents = SetCellContents(name, number);
        }
        else if (parsedContent is Formula formula)
        {
            all_dependents = SetCellContents(name, formula);
        }
        else
        {
            all_dependents = SetCellContents(name, (string)parsedContent);
        }

        this.Changed = true;
        foreach (string s in all_dependents)
        {
            if (cells.ContainsKey(s))
            {
                cells[s].RecalculateCell(LookUpValue);
            }
        }

        return all_dependents;
    }

    /// <summary>
    /// Parses the content string and returns the appropriate cell content object (string, double, or formula).
    /// </summary>
    /// <param name="content">The content string to parse.</param>
    /// <returns>The parsed content as an object (string, double, or formula).</returns>
    private static object ParseContent(string content)
    {
        if (double.TryParse(content, out double number))
        {
            return number;
        }
        else if (content.StartsWith("="))
        {
            return new Formula(content.Substring(1));
        }
        else
        {
            return content;
        }
    }

    /// <summary>
    ///  Set the contents of the named cell to the given number.
    /// </summary>
    ///
    /// <exception cref="InvalidNameException">
    ///   If the name is invalid, throw an InvalidNameException.
    /// </exception>
    ///
    /// <param name="name"> The name of the cell. </param>
    /// <param name="number"> The new content of the cell. </param>
    /// <returns>
    ///   <para>
    ///     This method returns an ordered list consisting of the passed in name
    ///     followed by the names of all other cells whose value depends, directly
    ///     or indirectly, on the named cell.
    ///   </para>
    ///   <para>
    ///     The order must correspond to a valid dependency ordering for recomputing
    ///     all of the cells, i.e., if you re-evaluate each cell in the order of the list,
    ///     the overall spreadsheet will be correctly updated.
    ///   </para>
    ///   <para>
    ///     For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
    ///     list [A1, B1, C1] is returned, i.e., A1 was changed, so then A1 must be
    ///     evaluated, followed by B1 re-evaluated, followed by C1 re-evaluated.
    ///   </para>
    /// </returns>
    private IList<string> SetCellContents(string name, double number)
    {
        name = name.ToUpper();
        this.ValidateCellName(name);

        Cell cell = new Cell(number);
        if (this.cells.ContainsKey(name))
        {
            this.cells[name] = cell;
        }
        else
        {
            this.cells.Add(name, cell);
        }

        // replace the dependents of 'name' in the dependency graph with an empty hash set
        this.dg.ReplaceDependees(name, new HashSet<string>());

        // recalculate at end
        IList<string> all_dependees = new List<string>(this.GetCellsToRecalculate(name));
        return all_dependees;
    }

    /// <summary>
    ///   The contents of the named cell becomes the given text.
    /// </summary>
    ///
    /// <exception cref="InvalidNameException">
    ///   If the name is invalid, throw an InvalidNameException.
    /// </exception>
    /// <param name="name"> The name of the cell. </param>
    /// <param name="text"> The new content of the cell. </param>
    /// <returns>
    ///   The same list as defined in <see cref="SetCellContents(string, double)"/>.
    /// </returns>
    private IList<string> SetCellContents(string name, string text)
    {
        name = name.ToUpper();
        this.ValidateCellName(name);
        Cell cell = new Cell(text);
        if (text != string.Empty)
        {
            if (this.cells.ContainsKey(name))
            {
                this.cells[name] = cell;
            }
            else
            {
                this.cells.Add(name, cell);
            }
        }
        else
        {
            this.cells.Remove(name);
        }

        this.dg.ReplaceDependees(name, new HashSet<string>());

        // recalculate at end
        IList<string> all_dependees = new List<string>(this.GetCellsToRecalculate(name));
        return all_dependees;
    }

    /// <summary>
    ///   Set the contents of the named cell to the given formula.
    /// </summary>
    /// <exception cref="InvalidNameException">
    ///   If the name is invalid, throw an InvalidNameException.
    /// </exception>
    /// <exception cref="CircularException">
    ///   <para>
    ///     If changing the contents of the named cell to be the formula would
    ///     cause a circular dependency, throw a CircularException.
    ///   </para>
    ///   <para>
    ///     No change is made to the spreadsheet.
    ///   </para>
    /// </exception>
    /// <param name="name"> The name of the cell. </param>
    /// <param name="formula"> The new content of the cell. </param>
    /// <returns>
    ///   The same list as defined in <see cref="SetCellContents(string, double)"/>.
    /// </returns>
    private IList<string> SetCellContents(string name, Formula formula)
    {
        name = name.ToUpper();
        this.ValidateCellName(name);
        IEnumerable<string> oldDependees = this.dg.GetDependees(name);

        // replace the dependents of 'name' in the dependency graph with the variables in formula
        // check if the new depdendency graph creates a circular reference
        try
        {
            // replace the dependents of 'name' in the dependency graph with the variables in formula
            this.dg.ReplaceDependees(name, formula.GetVariables());

            // if there is no exception
            IList<string> all_dependees = new List<string>(this.GetCellsToRecalculate(name));

            // create a new cell
            Cell cell = new Cell(formula, this.LookUpValue);
            if (this.cells.ContainsKey(name))
            {
                this.cells[name] = cell;
            }
            else
            {
                this.cells.Add(name, cell);
            }

            return all_dependees;
        }
        catch (CircularException)
        {
            this.dg.ReplaceDependees(name, oldDependees);
            throw new CircularException();
        }
    }

    /// <summary>
    ///   Returns an enumeration, without duplicates, of the names of all cells whose
    ///   values depend directly on the value of the named cell.
    /// </summary>
    /// <param name="name"> This <b>MUST</b> be a valid name.  </param>
    /// <returns>
    ///   <para>
    ///     Returns an enumeration, without duplicates, of the names of all cells
    ///     that contain formulas containing name.
    ///   </para>
    ///   <para>For example, suppose that: </para>
    ///   <list type="bullet">
    ///      <item>A1 contains 3</item>
    ///      <item>B1 contains the formula A1 * A1</item>
    ///      <item>C1 contains the formula B1 + A1</item>
    ///      <item>D1 contains the formula B1 - C1</item>
    ///   </list>
    ///   <para> The direct dependents of A1 are B1 and C1. </para>
    /// </returns>
    private IEnumerable<string> GetDirectDependents(string name)
    {
        this.ValidateCellName(name);
        return this.dg.GetDependents(name);
    }

    /// <summary>
    ///   <para>
    ///     This method is implemented for you, but makes use of your GetDirectDependents.
    ///   </para>
    ///   <para>
    ///     Returns an enumeration of the names of all cells whose values must
    ///     be recalculated, assuming that the contents of the cell referred
    ///     to by name has changed.  The cell names are enumerated in an order
    ///     in which the calculations should be done.
    ///   </para>
    ///   <exception cref="CircularException">
    ///     If the cell referred to by name is involved in a circular dependency,
    ///     throws a CircularException.
    ///   </exception>
    ///   <para>
    ///     For example, suppose that:
    ///   </para>
    ///   <list type="number">
    ///     <item>
    ///       A1 contains 5
    ///     </item>
    ///     <item>
    ///       B1 contains the formula A1 + 2.
    ///     </item>
    ///     <item>
    ///       C1 contains the formula A1 + B1.
    ///     </item>
    ///     <item>
    ///       D1 contains the formula A1 * 7.
    ///     </item>
    ///     <item>
    ///       E1 contains 15
    ///     </item>
    ///   </list>
    ///   <para>
    ///     If A1 has changed, then A1, B1, C1, and D1 must be recalculated,
    ///     and they must be recalculated in an order which has A1 first, and B1 before C1
    ///     (there are multiple such valid orders).
    ///     The method will produce one of those enumerations.
    ///   </para>
    /// </summary>
    /// <param name="name"> The name of the cell.  Requires that name be a valid cell name.</param>
    /// <returns>
    ///    Returns an enumeration of the names of all cells whose values must
    ///    be recalculated.
    /// </returns>
    private IEnumerable<string> GetCellsToRecalculate(string name)
    {
        LinkedList<string> changed = new();
        HashSet<string> visited = new();
        this.Visit(name, name, visited, changed);
        return changed;
    }

    /// <summary>
    ///   A helper for the GetCellsToRecalculate method.
    /// </summary>
    private void Visit(string start, string name, ISet<string> visited, LinkedList<string> changed)
    {
        visited.Add(name);
        foreach (string dependent in this.GetDirectDependents(name))
        {
            if (dependent.Equals(start))
            {
                throw new CircularException();
            }
            else if (!visited.Contains(dependent))
            {
                this.Visit(start, dependent, visited, changed);
            }
        }

        changed.AddFirst(name);
    }

    /// <summary>
    /// Validates the cell name provided and throws an <see cref="InvalidNameException"/>
    /// if the name is invalid.
    /// </summary>
    /// <param name="name">The name of the cell to validate.</param>
    /// <exception cref="InvalidNameException">
    /// Thrown when the provided cell name does not meet the valid naming criteria.
    /// </exception>
    private void ValidateCellName(string name)
    {
        if (!Regex.IsMatch(name, @"^[a-zA-Z]+\d+$"))
        {
            throw new InvalidNameException();
        }
    }

    /// <summary>
    /// Helper method retrieves the numeric value of a specified cell from the dependency graph.
    /// </summary>
    /// <param name="s">The key of the cell to look up.</param>
    /// <returns>
    /// The numeric value of the specified cell as a <see cref="double"/>.
    /// </returns>
    /// <exception cref="ArgumentException">
    /// Thrown when the specified cell does not exist in the dependency graph or
    /// if the cell's value is not of type <see cref="double"/>.
    /// </exception>
    private double LookUpValue(string s)
    {
        if (cells.ContainsKey(s))
        {
            if (cells[s].Value is double)
            {
                return (double)cells[s].Value;
            }
            else
            {
                throw new ArgumentException();
            }
        }

        throw new ArgumentException();
    }

    /// <summary>
    /// Represents a cell in a spreadsheet that can hold different types of content,
    /// including strings, doubles, and formulas. The content type is determined
    /// by the constructor used during instantiation.
    /// </summary>
    private class Cell
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class.
        /// Initializes a new cell with a text (string) as its content.
        /// Sets the content type to "string" and initializes an empty set for formula references.
        /// </summary>
        /// <param name="text">The string content to store in the cell.</param>
        public Cell(string text)
        {
            this.Contents = text;
            this.Value = text;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class.
        /// Initializes a new cell with a numeric value (double) as its content.
        /// Sets the content type to "double" and initializes an empty set for formula references.
        /// </summary>
        /// <param name="number">The double value to store in the cell.</param>
        public Cell(double number)
        {
            this.Contents = number;
            this.Value = number;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class.
        /// Initializes a new cell with a formula as its content.
        /// Sets the content type to "formula" and extracts the variables referenced by the formula.
        /// </summary>
        /// <param name="formula">The formula to store in the cell.</param>
        /// <param name="lookup">The method to find the value of the cell.</param>
        public Cell(Formula formula, Func<string, double> lookup)
        {
            Contents = formula;
            Value = formula.Evaluate(new Lookup(lookup));
        }

        /// <summary>
        /// Gets the content stored in the cell. This can be a string, a double, or a formula.
        /// </summary>
        public object Contents { get; }

        /// <summary>
        /// Gets the value stored in the cell. This can be a string, a double, or a formula.
        /// </summary>
        public object Value { get; private set; }

        /// <summary>
        /// Recalculates the value of the cell if its content is a formula.
        /// This method evaluates the formula using the provided lookup function to fetch the values of dependent cells.
        /// If the content is not a formula, no action is taken.
        /// </summary>
        /// <param name="lookup">
        /// A function that takes a cell name (string) and returns its numeric value (double).
        /// This lookup function is used to evaluate any variables referenced in the formula.
        /// </param>
        public void RecalculateCell(Func<string, double> lookup)
        {
            if (this.Contents is Formula)
            {
                Formula same = (Formula)this.Contents;
                this.Value = same.Evaluate(new Lookup(lookup));
            }
        }
    }
}