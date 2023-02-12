#!/bin/bash

# Define your GitHub username
username=<your-github-username>

# Get the list of all repositories for the user
repos=$(curl -s https://api.github.com/users/$username/repos?per_page=1000 | jq '.[].name' | tr -d '"')

# Backup each repository to the local system
for repo in $repos; do
    if [ -d "$repo" ]; then
        # Update the existing repository if it already exists
        (cd $repo && git fetch --all && git pull)
    else
        # Clone the repository if it doesn't already exist
        git clone --recursive https://github.com/$username/$repo.git
    fi
    # Compress the repository into a tar archive
    # tar czf $repo.tar.gz $repo
done

