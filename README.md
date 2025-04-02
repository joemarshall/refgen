# Installation

Install with pip, i.e. clone the git to `some_dir/refgen` and run `pip some_dir/refgen`

# Usage

You need a docx file called `reference_template.docx` in the current folder, then type
`refgen`.

Add a reference letter, then once you've done that, select it on the main screen and hit enter (or 'generate') and
you should have a `.docx` (whatever platform) and a `.pdf` (windows only) in the current folder. If anyone wants to submit a PR for
mac/linux support you're welcome!

# Template format

The template document is a normal word docx file, with tags embedded directly into the text. The tags use Jinja tags which mean you can do things like conditionals depending on whether the student is currently studying or
has finished.

Tags are:

|tag|meaning|
|------|------|
|student_name|The name of the student|
|ref_date| The date on the reference letter|
|how_known| A sentence to say how you know the student |
|start_date|The start date of the student's studies |
|end_date|The end date of the student's studies|
|has_end|If this is false, the student is still studying|
|recommendation_text|The text recommending the student|
|target|A target job / university / whatever the student is applying for|

I use Jinja rather than word templates because it makes it easier to do e.g. conditionals.

```Jinja
 This is to confirm that they {% if has_end %} studied Computer Science
 at the University of Nottingham from {{ start_date }} until {{ end_date }}
{% else %} have been a student of Computer Science at the University of Nottingham
 since {{ start_date }}{% endif %}
```
