SAURABIX - HOW TO RUN (free)

1. Install Python from python.org  (tick "Add Python to PATH" during install)
2. Get a free API key: go to aistudio.google.com -> sign in -> "Get API key"
3. Open Command Prompt / terminal inside this saurabix folder and run:
      pip install -r requirements.txt
4. Start the app:
      streamlit run app.py
   Your browser opens. Paste the API key in the sidebar. Done!

IF IT SHOWS A MODEL ERROR: open app.py and change the MODEL line at the top
to the model name shown in Google AI Studio.

NEVER upload your API key to GitHub or share it.
