@echo off

cd target
rem java -cp "lib/*;config;ListComparator-1.0-SNAPSHOT.jar" fr.lsmbo.Main test-classes\testFile.xlsx 1 B A C,D,E,F,G ../output.xlsx
rem java -cp "lib/*;config;ListComparator-1.0-SNAPSHOT.jar" fr.lsmbo.Main test-classes\small.xlsx 0 B A C,D,E,F,G ../output.xlsx
java -jar ListComparator-1.0-SNAPSHOT-jar-with-dependencies.jar test-classes\small.xlsx 0 B A C,D,E,F,G ../output.xlsx

cd ..

rem pause
