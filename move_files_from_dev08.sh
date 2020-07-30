#dir=$1
loose=$1
loankit_process_id=$2 
infynity_process_id=$3

echo -e "Loose: ${loose}|loankit_process_id: ${loankit_process_id}|infynity_process_id:${infynity_process_id}\n"
echo -e "./move_files_from_dev08.sh ${loose} ${loankit_process_id} ${infynity_process_id}\n"
#echo -e "scp -r -P 4422 qaisar@203.98.82.13:/var/www/mystro.com/data/commission_loankit/rcti/pre_processed/${infynity_process_id} .\n "
#scp -r -P 4422 qaisar@203.98.82.13:/var/www/mystro.com/data/commission_loankit/rcti/pre_processed/${infynity_process_id} /Users/qaisar/Desktop/rcti_comparison_scripts/commission-comparer-infynity/inputs/downloads/infynity/
echo -e "python /Users/qaisar/Desktop/rcti_comparison_scripts/commission-comparer-infynity/scripts/move_infynity_files_to_test_dir.py ${infynity_process_id}\n"
python /Users/qaisar/Desktop/rcti_comparison_scripts/commission-comparer-infynity/scripts/move_infynity_files_to_test_dir.py ${infynity_process_id}
echo -e "python /Users/qaisar/Desktop/rcti_comparison_scripts/commission-comparer-infynity/scripts/move_loankit_files_to_test_dir.py ${loankit_process_id}\n"
python /Users/qaisar/Desktop/rcti_comparison_scripts/commission-comparer-infynity/scripts/move_loankit_files_to_test_dir.py ${loankit_process_id}
echo -e "python /Users/qaisar/Desktop/rcti_comparison_scripts/commission-comparer-infynity/cli.py ${loose} ${loankit_process_id} ${infynity_process_id}\n"
python /Users/qaisar/Desktop/rcti_comparison_scripts/commission-comparer-infynity/cli.py ${loose} ${loankit_process_id} ${infynity_process_id}
