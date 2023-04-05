# Changelog

## 2.5.0 (2020-04-05)

- Ensure that we escape the first underscore character in plaintext strings that match the format for Excel escape sequences. 

## 2.4.0 (2020-06-27)

- Allow writing worksheets without a block using add\_worksheet (#42, #45)
- Deprecate calling add\_worksheet with a block, use write\_worksheet instead (#45)
- Relax rubyzip development dependency to allow current version (#46)

## 2.3.0 (2019-11-27)

- Speed up date / time conversion to OA format (#39)

## 2.2.0 (2019-11-27)

- Allow usage with zip\_tricks 5.x gem (#38)

## 2.1.0 (2018-07-21)

- New `:columns` option, allowing column widths to be specified (#25)
- Fix compatibility with `ruby --enable-frozen-string-literal` (#27)
- Support giving the worksheet name as an option to write\_worksheet (#28)

## 2.0.1 (2018-03-11)

- Rescue gracefully from invalid dates with auto-format (#22)
- Remove unused ZipTricksFibers IO wrapper (#24)

## 2.0.0 (2017-10-31)

- Replace RubyZip with ZipTricks as default compressor (#16)
- Drop support for Ruby < 2.1.0 (required for zip\_tricks gem)
- Deprecate :io\_wrapper option, you can now pass wrapper instances (#20)

## 1.3.2 (2017-10-30)

- Fix circular require in workbook
- Fix wrong escaping of extended Unicode characters

## 1.3.1 (2017-10-30)

- Restore stringio require in workbook

## 1.3.0 (2017-10-30)

- Drop rubyzip buffering workarounds, require rubyzip >= 1.2.0 (#17)
- Drop Ruby 1.9.1 compatibility (rubyzip 1.2 requires ruby >= 1.9.2)
- Refactor IO wrappers (#18)

## 1.2.0 (2017-10-30)

- Add support for customizing default font (#17)

## 1.1.0 (2017-10-23)

- Add support for boolean values (#13)

## 1.0.1 (2017-10-22)

- Fix writing unnamed worksheets with options

## 1.0.0 (2017-10-22)

- Don't close IO objects passed to Workbook constructor (#12)

## 0.3.1 (2017-10-22)

- Escape invalid XML 1.0 characters (#11)

## 0.3.0 (2017-07-12)

- Add support for auto-formatting (#8)

## 0.2.0 (2017-02-20)

- Ruby 2.4 compatibility
- Misc bug fixes

## 0.1.0 (2015-10-17)

- Initial release
