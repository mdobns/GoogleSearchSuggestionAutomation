from selenium import webdriver
from selenium.webdriver.common.by import By
import time

class GoogleSearchSuggestions:
    def __init__(self):
        self.driver = webdriver.Firefox()
        self.driver.get("http://www.google.com")

    def search(self, search_query):
        search_bar = self.driver.find_element(By.NAME, "q")
        search_bar.clear()  # Clear the search box
        search_bar.send_keys(search_query)
        time.sleep(2)

    def get_suggestions(self):
        search_text_google_suggestions = self.driver.find_elements(
            By.XPATH, "//ul[@role=\"listbox\"]/li")
        suggestion_list = [suggestion.text for suggestion in search_text_google_suggestions]
        return suggestion_list

    def get_max_min_suggestions(self, search_query):
        self.search(search_query)
        suggestion_list = self.get_suggestions()

        if suggestion_list:
            maximum_suggestion = max(suggestion_list, key=len)
            minimum_suggestion = min(suggestion_list, key=len)
            return maximum_suggestion, minimum_suggestion
        else:
            return None, None

    def close_browser(self):
        self.driver.quit()
