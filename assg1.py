import pandas as pd

# Load the raw data from the Excel file
file_path = 'C:\Users\user\Downloads/Football Data Test Task.xlsx'
raw_data = pd.read_excel(file_path, sheet_name='Raw Data')

# Assuming the raw data has columns for match results
# Example structure: ['Home Team', 'Away Team', 'Home Wins', 'Away Wins', ...]
# Adjust the column names based on your actual data

# Function to calculate total wins for a team in the last N matches
def calculate_wins(team_name, matches, N):
    # Filter matches for the specific team
    team_matches = matches[(matches['Home Team'] == team_name) | (matches['Away Team'] == team_name)]
    
    # Get the last N matches
    last_n_matches = team_matches.tail(N)
    
    # Calculate total wins
    total_wins = 0
    for _, match in last_n_matches.iterrows():
        if match['Home Team'] == team_name:
            total_wins += match['Home Wins']  # Assuming 'Home Wins' is the column with win counts
        else:
            total_wins += match['Away Wins']  # Assuming 'Away Wins' is the column with win counts
            
    return total_wins

# Create a new DataFrame for manipulated data
manipulated_data = pd.DataFrame()

# Iterate through each row in the raw data
for index, row in raw_data.iterrows():
    home_team = row['Home Team']  # Assuming 'Home Team' is the column name
    away_team = row['Away Team']  # Assuming 'Away Team' is the column name
    
    # Calculate wins for last 5 matches
    home_wins_last_5 = calculate_wins(home_team, raw_data, 5)
    away_wins_last_5 = calculate_wins(away_team, raw_data, 5)
    
    # Calculate wins for last 15 matches
    home_wins_last_15 = calculate_wins(home_team, raw_data, 15)
    away_wins_last_15 = calculate_wins(away_team, raw_data, 15)
    
    # Calculate wins for last 38 matches
    home_wins_last_38 = calculate_wins(home_team, raw_data, 38)
    away_wins_last_38 = calculate_wins(away_team, raw_data, 38)
    
    # Append results to manipulated data
    manipulated_data = manipulated_data.append({
        'Home Team': home_team,
        'Away Team': away_team,
        'Home Wins Last 5': home_wins_last_5,
        'Away Wins Last 5': away_wins_last_5,
        'Home Wins Last 15': home_wins_last_15,
        'Away Wins Last 15': away_wins_last_15,
        'Home Wins Last 38': home_wins_last_38,
        'Away Wins Last 38': away_wins_last_38
    }, ignore_index=True)

# Save the manipulated data to a new Excel file
manipulated_data.to_excel('Manipulated_Data.xlsx', index=False)

print("Data manipulation complete. Results saved to 'Manipulated_Data.xlsx'.")