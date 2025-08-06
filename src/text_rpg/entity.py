"""
Entity

An entity is either the player or enemies. Anything that has health and can do damage
"""

class Entity:
    health = 0
    level = 0
    
class Hero(Entity):
    mana = 0
