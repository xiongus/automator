#!/bin/bash

# Define the directory that contains your Git projects
root_dir=<directory-containing-your-git-projects>

# Loop through all subdirectories of the root directory
for dir in $(find $root_dir -type d -name '.git' | xargs -n1 dirname); do
    # Go to the Git repository
    cd $dir

    # Pull changes from remote repository
    git pull

    # Check if there are any changes to commit
    changes=$(git status --porcelain)
    if [ -n "$changes" ]; then
        # Commit changes
        git add .
        git commit -m "Automatic commit $(date)"

        # Push changes to remote repository
        git push
    fi

    # Return to the root directory
    cd $root_dir
done

