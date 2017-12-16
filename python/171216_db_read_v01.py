import psycopg2

try:
    conn = psycopg2.connect("dbname='DataCollecitonDB' user='postgres' host='localhost' password='kalman'")
    print "Connected Succesfully!"
except:
    print "I am unable to connect to the database"


