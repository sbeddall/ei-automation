## OpenPyXL Measurement Template Generator

Based off an input excel sheet, generates an `xlsx` as required for industry-specific use. 

### Usage

1. Install [Python](https://www.python.org/downloads/)
    * Make certain that the `Install to Path` checkbox available during the installation is **checked**. 
3. To generate an excel sheet, update the template excel (`User Input` sheet) to the station numbers you want. 
4. Open up a command line window, `cd` to the directory containing both the template and the python script.
5. In the command line, run the following command `python generate_excel.py`
	* As a **one-time** step before running this script for the first time, you need to run the following command on the command line `pip install openpyxl`
6. Look in the folder alongside the script. There will be a `generated_xxxx` file that contains the templates.

### Results

![generated excel](example.png "Generated Excel")

### Compatibility

Python 3.7+. My users install latest Python as required. Not worth to maintain back-compat.

<!-- Google Analytics -->
<script>
(function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){
(i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o),
m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)
})(window,document,'script','https://www.google-analytics.com/analytics.js','ga');

ga('create', 'UA-133754601-1', 'auto');
ga('send', 'pageview');
</script>

<!-- Global site tag (gtag.js) - Google Analytics -->
<script async src="https://www.googletagmanager.com/gtag/js?id=UA-133754601-1"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());

  gtag('config', 'UA-133754601-1');
</script>

<!-- End Google Analytics -->