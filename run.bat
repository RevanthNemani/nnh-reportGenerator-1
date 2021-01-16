if exist outputs Del /F /Q "outputs"
pause
mkdir outputs
Rscript.exe "reportGenerator.R"
