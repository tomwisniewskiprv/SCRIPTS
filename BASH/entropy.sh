#!/bin/bash

# Script calculates Shannon entropy for two files and compares results.
#
# usage:
# entropy.sh <file1> <file2> 
# 
# author:
# Tomasz Wiśniewski gr B4
# Inforamtyka Zaoczna 2016/2017 Uniwersytet Śląski
# Rok I


# check for correct input
if [ ! -e $1 ] || [ ! -e $2 ] || [ "$#" -ne 2 ]; then
    echo "Wrong file names. Exiting."
    echo "usage:"
    echo "./entropy.sh <file1> <file2>"
    exit 1
fi

# create ASCII table with all printable chars
alfa=""
for i in {32..127}; do
    alfa+=`printf "\x$(printf %x $i)"`
done
alfa_len=${#alfa}

# hisogram arrays
hist_1=()
hist_2=()

# entropy records
entropy_1=0
entropy_2=0

# read first file input
echo "Reading first file."
fin=$1

declare -a words
words=(`cat "$fin"`)

# count and compare ASCII characters from file
for word in ${words[*]}; do
    word_len=${#word}

    for ((i=0 ; i<word_len ; i++)); do
        # single char from word 
        w=${word:$i:1}
       
        # compare with ASCII table 
        for ((c=0; c<$alfa_len; c++)); do
            ch=${alfa:$c:1}
            if [ "$w" == "$ch" ]; then
                hist_1[$c]+="="
                count=${hist_1[$c]}
                count=${#count}
            fi
        done
    done
done

# count characters in first file
sum_first_file=0
for char in ${hist_1[@]}; do
    len=$(echo ${#char})
    sum_first_file=$(($sum_first_file + $len))
done

# count entropy for first file
# note:
# echo "l(0.333)/l(2)*0.333" | bc -l

for char in ${hist_1[@]}; do
    chars=${#char}
    fraction=$(echo "scale=5; $chars/$sum_first_file" | bc)
    log=$(echo "l($fraction)/l(2)*$fraction" | bc -l)
    entropy_1=$(echo "$log + $entropy_1" | bc)
done

# first entropy value is ready
entropy_1=$(echo "$entropy_1 * -1" | bc)
echo "Entropy: $entropy_1"

# proceed with second file 
echo "Reading second file."
fin2=$2 

declare -a words
words=(`cat "$fin2"`)

# count and compare ASCII characters from file
for word in ${words[*]}; do
    word_len=${#word}

    for ((i=0 ; i<word_len ; i++)); do
        # single char from word 
        w=${word:$i:1}
       
        # compare with ASCII table 
        for ((c=0; c<$alfa_len; c++)); do
            ch=${alfa:$c:1}
            if [ "$w" == "$ch" ]; then
                hist_2[$c]+="="
                count=${hist_2[$c]}
                count=${#count}
            fi
        done
    done
done

# count characters in second file
sum_first_file=0
for char in ${hist_2[@]}; do
    len=$(echo ${#char})
    sum_first_file=$(($sum_first_file + $len))
done

# count entropy for second file
# note:
# echo "l(0.333)/l(2)*0.333" | bc -l

for char in ${hist_2[@]}; do
    chars=${#char}
    fraction=$(echo "scale=5; $chars/$sum_first_file" | bc)
    log=$(echo "l($fraction)/l(2)*$fraction" | bc -l)
    entropy_2=$(echo "$log + $entropy_2" | bc)
done

# second entropy value is ready
entropy_2=$(echo "$entropy_2 * -1" | bc)
echo "Entropy: $entropy_2"

# compare values
same_value=`echo "$entropy_1"'=='"$entropy_2" | bc -l`
cmp=`echo "$entropy_1"'>'"$entropy_2" | bc -l`

echo ""
echo "Result:"

if [ $same_value -ne 1 ]; then
    if [ $cmp -eq 0 ]; then
       echo "Entropy from file $fin2 is greater" 
    elif [ $cmp -eq 1 ]; then
        echo "Entropy from file $fin is greater."
    fi
else
    echo "Both files have exactly the same entropy."
fi

