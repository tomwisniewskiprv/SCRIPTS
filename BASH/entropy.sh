#!/bin/bash

# Script calculates Shannon entropy for two files and compares results.
#
# Tomasz Wiśniewski gr B4
# Inforamtyka Zaoczna 2016/2017 Uniwersytet Śląski
# Rok I


calculate_entropy(){
    words=($(cat "$1"))

    # count and compare ASCII characters from file
    for word in "${words[*]}"; do
        word_len=${#word}

        for ((i=0 ; i < "$word_len" ; i++)); do
            # single char from word 
            w=${word:$i:1}
           
            # compare with ASCII table 
            for ((c=0; c < "$alfa_len" ; c++)); do
                ch=${alfa:$c:1}
                if [ "'$w'" == "'$ch'" ]; then
                    hist[$c]+="="
                    count=${hist[$c]}
                    count=${#count}
                fi
            done
        done
    done

    # count characters in file
    sum=0
    for char in ${hist[@]}; do
        len=${#char}
        #len=$(echo ${#char})
        sum=$(($sum + $len))
    done

    # count entropy for file
    entropy=0
    for char in ${hist[@]}; do
        chars=${#char}
        fraction=$(echo "scale=5; $chars/$sum" | bc)
        log=$(echo "l($fraction)/l(2)*$fraction" | bc -l)
        entropy=$(echo "$log + $entropy" | bc -l)
    done

    # entropy value is ready
    entropy=$(echo "$entropy * -1" | bc)
    echo "$entropy"
}

# start
# check for correct input files
if [ ! -e $1 ] || [ ! -e $2 ] || [ "$#" -ne 2 ]; then
    echo "Wrong file names. Exiting."
    echo "usage:"
    echo "./entropy.sh <file1> <file2>"
    exit 1
fi

# create ASCII table with all printable chars
for i in {32..127}; do
    alfa+=$(printf "\x$(printf %x $i)")
done
alfa_len=${#alfa}

# read first file input
echo "Reading first file."
entropy_1=$(calculate_entropy $1)
echo "$entropy_1"

echo "Reading second file."
entropy_2=$(calculate_entropy $2)
echo "$entropy_2"

# result
# compare values
if [ "$(echo "$entropy_1 == $entropy_2" | bc)" -eq 1 ]; then
    echo "Both files have the same entropy."
elif [ "$(echo "$entropy_1 > $entropy_2" | bc)" -eq 1 ]; then
    echo "File $1 has greater entropy: $entropy_1"
elif [ "$(echo "$entropy_1 > $entropy_2" | bc)" -eq 0 ]; then
    echo "File $2 has greater entropy: $entropy_2"
fi


# the end
