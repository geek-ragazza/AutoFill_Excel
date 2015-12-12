pushd %~dp0\python-2.7.10
python Project_Evaluate_Excel\Search_History\Search_console.py 
echo "Hello from batch land"
echo %~dp0
popd
