import time 
import dateparser 
import pandas as pd
from selenium import webdriver
from datetime import date, datetime
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

class GoogleReviews:
    def __init__(self,search, file_name, path):
        self.search = search 
        self.name = file_name
        self.path = path
        self.driver = None
        self.all_reviews = None
        self.reviewer = []
        self.date = [] 
        self.rating = []
        self.review_description = []
        self.likes = []
        self.owner_response = []
        self.owner_response_date = []
        self.df = None 
        
        self.load_chrome()
        self.search_box()
        self.expand_reviews()
        self.scroll_down_to_end()
        self.get_names()
        self.get_ratings()
        self.get_review_date()
        self.get_review()
        self.get_likes()
        self.get_owner_response()
        self.get_owner_response_date()
        self.transform_dataframe()
        self.export_reviews_as_excel()
        self.driver.close() 
        
    def load_chrome(self):
        self.driver = webdriver.Chrome(executable_path=self.path)
        self.driver.get("https://www.google.com") # Load default Google page  
        time.sleep(2)
        self.driver.maximize_window() # Maximize window size 
        
    def search_box(self):
        search_box = self.driver.find_element_by_name("q")
        search_box.send_keys(self.search)
        search_box.submit() # Send keys
        
    def expand_reviews(self):
        expand_reviews = self.driver.find_element_by_partial_link_text("Google reviews") # Locate link 
        self.total_number_of_reviews = int(expand_reviews.get_attribute('textContent').split(" ")[0].replace(",", ""))
        expand_reviews.click() # Click to expand Google reviews 
    
    def scroll_down_to_end(self):
        self.all_reviews = WebDriverWait(self.driver, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.gws-localreviews__google-review')))
        while len(self.all_reviews) < self.total_number_of_reviews:
            self.driver.execute_script('arguments[0].scrollIntoView(true);', self.all_reviews[-1])
            self.all_reviews = self.driver.find_elements_by_css_selector('div.gws-localreviews__google-review') 
      
    def get_names(self):
        for review in self.all_reviews:
            try: 
                name_element = review.find_element_by_class_name("TSUbDb")
            except NoSuchElementException:
                pass
            finally:
                self.reviewer.append(name_element.text) 
        
    def get_ratings(self):
        for review in self.all_reviews:
            try:
                ratings_element = review.find_element_by_xpath("//*[@id='reviewSort']/div/div[2]/div[2]/div[1]/div[3]/div[1]/g-review-stars/span")
            except NoSuchElementException:
                pass
            finally:
                self.rating.append(ratings_element.get_attribute('aria-label').replace(",", "")) 
        
    def get_review_date(self):
        for review in self.all_reviews:
            try:
                date_element = review.find_element_by_class_name("dehysf").text.replace("a ", "1 ") 
            except NoSuchElementException:
                pass
            finally:
                self.date.append(dateparser.parse(date_element).date())
    
    def get_review(self):
        for review in self.all_reviews:
            try:
               full_text = review.find_element_by_css_selector('span.review-full-text')
            except: 
               full_text = review.find_element_by_class_name("Jtu6Td")
            finally:
               self.review_description.append(full_text.get_attribute('textContent'))
    
    def get_likes(self):
        for like in self.all_reviews:
            try:
                likes_element = like.find_element_by_css_selector('span[class = "QWOdjf"]')
                number = int(likes_element.get_attribute('textContent'))
                self.likes.append(number)
            except StaleElementReferenceException:
                pass
            except NoSuchElementException:
                self.likes.append(0)
    
    def get_owner_response(self):   
        for review in self.all_reviews:    
            try:
                review.find_element_by_css_selector("div.lororc").click()
            except Exception:
                pass
            try:
                response = review.find_element_by_css_selector('div.lororc > span:nth-child(3)').text
                self.owner_response.append(response)
            except NoSuchElementException:
                try:
                    response = review.find_element_by_css_selector('div.LfKETd').text.split('ago')[1]
                    self.owner_response.append(response)
                except NoSuchElementException:
                    self.owner_response.append('NA')
                
    def get_owner_response_date(self):
        for date in self.all_reviews:
            try:
                full_text_element = date.find_element_by_css_selector('span[class = "pi8uOe"')
                duration = dateparser.parse(full_text_element.get_attribute('textContent').replace("a ", "1 ")).date() 
                self.owner_response_date.append(duration)
            except StaleElementReferenceException:
                pass
            except NoSuchElementException:
                self.owner_response_date.append('NA')
    
    def transform_dataframe(self):
        self.df = pd.DataFrame(
            {'Name': self.reviewer,
             'Ratings': self.rating,
             'Review Date': self.date,
             'Review': self.review_description,
             'Likes': self.likes,
             'Owner Response': self.owner_response,
             'Response Date': self.owner_response_date
             })
        return self.df
    
    def export_reviews_as_excel(self):
        
        fname = self.name.replace(" ", "_")
        curr_date = datetime.today().strftime('%d%m%y')
        fformat = ".xlsx"
        full_fname = fname + "_" + curr_date + fformat
        
        self.df.to_excel(full_fname, index=True, index_label="unique_key")
        print(f'Exported as {full_fname}.')