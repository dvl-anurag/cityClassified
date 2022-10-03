package model

import (
	"go.mongodb.org/mongo-driver/bson/primitive"
)

type CityData struct {
	ID            primitive.ObjectID `bson:"_id,omitempty" json:"id,omitempty"`
	Title         string             `bson:"title,omitempty" json:"title,omitempty"`
	Name          string             `bson:"name,omitempty" json:"name,omitempty"`
	Address       string             `bson:"address, omitempty" json:"address,omitempty"`
	Latitude      float64            `bson:"latitude, omitempty" json:"latitude,omitempty"`
	Longitude     float64            `bson:"longitude, omitempty" json:"longitude,omitempty"`
	Website       string             `bson:"website, omitempty" json:"website,omitempty"`
	ContactNumber int64              `bson:"contact_number, omitempty" json:"contact_number,omitempty"`
	User          string             `bson:"user, omitempty" json:"user,omitempty"`
	City          string             `bson:"city, omitempty" json:"city,omitempty"`
	Country       string             `bson:"country, omitempty" json:"country,omitempty"`
	PinCode       int64              `bson:"pinCode, omitempty" json:"pinCode,omitempty"`
	UpdatedBy     string             `bson:"updated_by, omitempty" json:"updated_by,omitempty"`
	CategoriesId  primitive.ObjectID `bson:"categories_id,omitempty" json:"categories_id,omitempty"`
}

type Categories struct {
	ID       primitive.ObjectID `bson:"_id,omitempty" json:"id,omitempty"`
	Category string             `bson:"category, omitempty" json:"category,omitempty"`
}

type Search struct {
	Key   string `bson:"key,omitempty" json:"key,omitempty"`
	Value string `bson:"value,omitempty" json:"value,omitempty"`
}

type SearchBoth struct {
	City     string `bson:"city,omitempty" json:"city,omitempty"`
	Category string `bson:"category,omitempty" json:"category,omitempty"`
}
