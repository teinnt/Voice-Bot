void setup()
{
  Serial.begin(9600);
  pinMode(13, OUTPUT);
}

void loop()
{
  char id = Serial.read();
  
  if(id == 'A')
  {
    digitalWrite(13, HIGH);
  }
  
  if(id == 'B')
  {
    digitalWrite(13, LOW);
  }
}
