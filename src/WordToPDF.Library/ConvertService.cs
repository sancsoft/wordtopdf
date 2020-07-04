using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;

namespace WordToPDF.Library
{
    public class ConvertService
    {
        protected Dictionary<string, bool> _installedPrinters;
        protected Hashtable _options;
        protected bool _initialized;

        public ConvertService()
        {
            _initialized = false;
        }

        public void Initialize()
        {
            _installedPrinters = GetInstalledPrinters();
            if (_installedPrinters.Count == 0)
            {
                throw new Exception("No printers defined. A printer must be defined for output generation.");
            }
            _options = new Hashtable();
            _options["hidden"] = false;
            _options["markup"] = false;
            _options["readonly"] = false;
            _options["bookmarks"] = false;
            _options["print"] = true;
            _options["screen"] = false;
            _options["pdfa"] = false;
            _options["verbose"] = false;
            _options["excludeprops"] = false;
            _options["excludetags"] = false;
            _options["noquit"] = false;
            _options["merge"] = false;
            _options["template"] = "";
            _options["password"] = "";
            _options["printer"] = "";
            _options["fallback_printer"] = "";
            _options["working_dir"] = "";
            _options["has_working_dir"] = false;
            _options["excel_show_formulas"] = false;
            _options["excel_show_headings"] = false;
            _options["excel_auto_macros"] = false;
            _options["excel_template_macros"] = false;
            _options["excel_active_sheet"] = false;
            _options["excel_no_link_update"] = false;
            _options["excel_no_recalculate"] = false;
            _options["excel_max_rows"] = (int)0;
            _options["excel_worksheet"] = (int)0;
            _options["excel_delay"] = (int)0;
            _options["word_field_quick_update"] = false;
            _options["word_field_quick_update_safe"] = false;
            _options["word_no_field_update"] = false;
            _options["word_header_dist"] = (float)-1;
            _options["word_footer_dist"] = (float)-1;
            _options["word_max_pages"] = (int)0;
            _options["word_ref_fonts"] = false;
            _options["word_keep_history"] = false;
            _options["word_no_repair"] = false;
            _options["word_show_comments"] = false;
            _options["word_show_revs_comments"] = false;
            _options["word_show_format_changes"] = false;
            _options["word_show_hidden"] = false;
            _options["word_show_ink_annot"] = false;
            _options["word_show_ins_del"] = false;
            _options["word_markup_balloon"] = false;
            _options["word_show_all_markup"] = false;
            _options["word_fix_table_columns"] = false;
            _options["original_filename"] = "";
            _options["original_basename"] = "";
            _options["powerpoint_output"] = "";
            _options["pdf_page_mode"] = null;
            _options["pdf_layout"] = null;
            _options["pdf_merge"] = (int)MergeMode.None;
            _options["pdf_clean_meta"] = (int)MetaClean.None;
            _options["pdf_owner_pass"] = "";
            _options["pdf_user_pass"] = "";
            _options["pdf_restrict_annotation"] = false;
            _options["pdf_restrict_extraction"] = false;
            _options["pdf_restrict_assembly"] = false;
            _options["pdf_restrict_forms"] = false;
            _options["pdf_restrict_modify"] = false;
            _options["pdf_restrict_print"] = false;
            _options["pdf_restrict_annotation"] = false;
            _options["pdf_restrict_accessibility_extraction"] = false;
            _options["pdf_restrict_full_quality"] = false;

            _initialized = true;
        }

        public void SetOption(string optionName, bool value)
        {
            _options[optionName] = value;
        }
        public void SetOption(string optionName, int value)
        {
            _options[optionName] = value;
        }
        public void SetOption(string optionName, string value)
        {
            _options[optionName] = value;
        }

        public bool GetOptionBool(string optionName)
        {
            return (bool)_options[optionName];
        }

        public int GetOptionInt(string optionName)
        {
            return (int)_options[optionName];
        }

        public string GetOptionBoolean(string optionName)
        {
            return (string)_options[optionName];
        }

        private static Dictionary<string, bool> GetInstalledPrinters()
        {
            Dictionary<string, bool> printers = new Dictionary<string, bool>();
            try
            {
                foreach (string name in PrinterSettings.InstalledPrinters)
                {
                    printers[name.ToLowerInvariant()] = true;
                }
            }
            catch (Exception) { }
            return printers;
        }

        public int Convert(string inputFile, ref string outputFile)
        {
            if (!_initialized)
            {
                throw new Exception("ConverService is not initialized");
            }

            // if no output is provided, use the source file and change to a PDF extension
            if (string.IsNullOrEmpty(outputFile))
            {
                outputFile = Path.ChangeExtension(inputFile, "pdf");
            }
            else
            {
                // if the outputFile spec is a directory, put the file in the directory with same name and new extension
                if (Directory.Exists(outputFile))
                {
                    outputFile = Path.Combine(outputFile, Path.GetFileNameWithoutExtension(inputFile) + ".pdf");
                }
            }

            // confirm that the input file exists and convert to component path elements in options
            FileInfo inputInfo;
            try
            {
                inputInfo = new FileInfo(inputFile);
            }
            catch
            {
                inputInfo = null;
            }
            if (inputInfo == null || !inputInfo.Exists)
            {
                throw new Exception("Input file not found");
            }
            inputFile = inputInfo.FullName;
            _options["original_filename"] = inputInfo.Name;
            _options["original_basename"] = inputInfo.Name.Substring(0, inputInfo.Name.Length - inputInfo.Extension.Length);

            // handle the output file existing or path existing
            FileInfo outputInfo = new FileInfo(outputFile);
            // Remove the destination unless we're doing a PDF merge
            if (outputInfo != null)
            {
                outputFile = outputInfo.FullName;
                if (outputInfo.Exists)
                {
                    if ((MergeMode)_options["pdf_merge"] == MergeMode.None)
                    {
                        // We are not merging, so delete the final destination
                        System.IO.File.Delete(outputInfo.FullName);
                    }
                    else
                    {
                        // We are merging, so make a temporary file
                        outputFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".pdf";
                    }
                }
                else
                {
                    // If there is no current output, no need to merge
                    _options["pdf_merge"] = MergeMode.None;
                }
            }
            else
            {
                throw new Exception("Unable to determine outputFile location");
            }
            if (!System.IO.Directory.Exists(outputInfo.DirectoryName))
            {
                throw new Exception("Output directory does not exist");
            }

            // actually do the conversion, finally
            int converted = WordConverter.Convert(inputFile, outputFile, _options);

            return converted;
        }

    }
}
