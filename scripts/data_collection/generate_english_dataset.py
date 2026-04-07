import json
import itertools
import random
import csv
import os

# Load taxonomy
taxonomy_path = os.path.join(os.path.dirname(__file__), '../../data/knowledge_base/taxonomy.json')
with open(taxonomy_path, 'r', encoding='utf-8') as f:
    taxonomy = json.load(f)

intents = [intent['name'] for intent in taxonomy['intents']]

# Define templates for each intent
# We'll use placeholders for entities
PRODUCTS = ["kurta", "saree", "t-shirt", "shoes", "jeans", "watch", "perfume", "laptop", "mobile", "bag"]
SIZES = ["S", "M", "L", "XL", "XXL", "7", "8", "9", "10", "free size"]
CITIES = ["Mumbai", "Delhi", "Bangalore", "Chennai", "Kolkata", "Hyderabad", "Pune", "Ahmedabad", "Surat", "Jaipur"]
DATES = ["today", "tomorrow", "yesterday", "Monday", "Friday", "next week", "2 days ago", "10th October", "15th August"]

TEMPLATES = {
    "track_order": [
        "Where is my order?",
        "Can you track my order?",
        "What is the status of my order {order_id}?",
        "When will order {order_id} arrive?",
        "Any update on my delivery?",
        "Has my {product} been shipped?",
        "I want to track order {order_id}.",
        "Give me the tracking link for my {product}.",
        "Is my order out for delivery?",
        "Why is my order delayed?"
    ],
    "cancel_order": [
        "I want to cancel my order.",
        "Cancel order {order_id}.",
        "Please abort my delivery for {product}.",
        "Can I cancel the order I placed {date}?",
        "Cancel the {product} I bought.",
        "I changed my mind, cancel order {order_id}.",
        "Stop the shipment for {order_id}.",
        "I don't need the {product} anymore, cancel.",
        "How do I cancel my recent order?",
        "Cancel this order immediately."
    ],
    "return_product": [
        "I want to return this {product}.",
        "How can I return order {order_id}?",
        "This {product} is defective, I need a return.",
        "Take back my order.",
        "What is your return policy?",
        "I don't like the {product}, arrange for a return.",
        "Can I return the {product} I bought {date}?",
        "Schedule a pickup for returning my order.",
        "I want to send the {product} back.",
        "The item is damaged, I need to return it."
    ],
    "refund_status": [
        "When will I get my refund?",
        "Where is my money for order {order_id}?",
        "Refund status for my returned {product}?",
        "I haven't received my refund yet.",
        "How many days for the refund to process?",
        "Has the refund been initiated?",
        "Check my refund status for {order_id}.",
        "I am waiting for my refund since {date}.",
        "When will the amount reflect in my bank account?",
        "Please credit my refund."
    ],
    "payment_issue": [
        "Money deducted but order not placed.",
        "Payment failed but amount was debited.",
        "I was charged twice for order {order_id}.",
        "My payment is stuck.",
        "Payment issue with my recent transaction.",
        "The app says payment failed, but my bank says successful.",
        "I paid for the {product} but didn't get confirmation.",
        "Why is my payment pending?",
        "Can you check if my payment went through?",
        "Amount deducted from my account erroneously."
    ],
    "product_availability": [
        "Is the {product} available in size {size}?",
        "When will the {product} be back in stock?",
        "Do you have this in stock?",
        "Is this item out of stock?",
        "Can you notify me when {product} is available?",
        "I want this {product}. Is it available?",
        "Check if size {size} is there for {product}.",
        "Are there any units left for {product}?",
        "Show me {product}s in stock.",
        "Availability of {product} please."
    ],
    "shipping_cost": [
        "What is the delivery charge?",
        "How much is shipping to {city}?",
        "Is shipping free?",
        "Why are you charging for delivery?",
        "Delivery cost for {product}?",
        "Do I have to pay shipping for order {order_id}?",
        "What are the shipping fees?",
        "Can I get free delivery on {product}?",
        "Shipping charges are too high.",
        "How much do you charge for shipping?"
    ],
    "delivery_time": [
        "When will my order {order_id} be delivered?",
        "How many days to deliver to {city}?",
        "Delivery time for {product}?",
        "Will I receive my order by {date}?",
        "What is the estimated delivery time?",
        "How fast can you deliver to {city}?",
        "Can I get it delivered {date}?",
        "When is the expected delivery?",
        "Is express delivery available?",
        "How long does shipping take?"
    ],
    "change_address": [
        "I want to change my delivery address.",
        "Update the address for order {order_id}.",
        "Can I change shipping address to {city}?",
        "I put the wrong address, please change it.",
        "Modify the delivery location.",
        "Ship the {product} to my new address.",
        "I need to update my shipping details.",
        "Change my address for order {order_id}.",
        "Deliver this to {city} instead.",
        "How to edit delivery address?"
    ],
    "exchange_product": [
        "I want to exchange this {product}.",
        "Can I exchange size {size} for another size?",
        "Exchange my order {order_id}.",
        "The {product} doesn't fit, want an exchange.",
        "Can I get a {product} in {size} instead?",
        "I need a smaller size for my {product}.",
        "Exchange POLICY?",
        "Replace this {product} with a larger one.",
        "I want to swap this item.",
        "How do I apply for an exchange?"
    ],
    "coupon_apply": [
        "My coupon code {coupon_code} is not working.",
        "Why can't I apply this discount?",
        "Coupon {coupon_code} says invalid.",
        "How to use promo code?",
        "I forgot to apply my coupon {coupon_code}.",
        "The discount code {coupon_code} is showing an error.",
        "Can I apply coupon on {product}?",
        "Promo code {coupon_code} is not accepted.",
        "My voucher is not working.",
        "Why isn't the {coupon_code} applying to my cart?"
    ],
    "account_login": [
        "I can't log into my account.",
        "My account is locked.",
        "I forgot my password.",
        "How do I reset my password?",
        "Account {account_id} is inaccessible.",
        "Unable to sign in.",
        "My login is failing.",
        "Unlock my profile please.",
        "I am not getting the OTP to login.",
        "Help me recover my account."
    ],
    "order_modify": [
        "Can I add {product} to my current order?",
        "I want to modify order {order_id}.",
        "Change the quantity to {quantity}.",
        "Can I change the attributes of order {order_id}?",
        "Modify my order items.",
        "I ordered {quantity} but I want more.",
        "Remove {product} from order {order_id}.",
        "Update my cart for order {order_id}.",
        "Can I alter my recent order?",
        "I want to change the items I bought."
    ],
    "store_hours": [
        "What time does your store open?",
        "Are you open {date}?",
        "Store timings in {city}?",
        "When do you close today?",
        "What are your business hours?",
        "Is the {city} branch open on Sunday?",
        "Tell me the operating hours.",
        "Until what time are you open?",
        "Store opening schedule?",
        "Are you currently open?"
    ],
    "complaint_register": [
        "I want to file a complaint.",
        "Register a complaint regarding order {order_id}.",
        "I am very disappointed with the {product}.",
        "Where can I register my grievance?",
        "I have a formal complaint to make.",
        "The service for order {order_id} was terrible.",
        "I need to report an issue with the delivery.",
        "Complaint about the courier in {city}.",
        "Take my complaint.",
        "Your customer service is bad, I want to complain."
    ],
    "product_review": [
        "How do I review this {product}?",
        "I want to leave a rating.",
        "Where can I write a product review?",
        "Can I upload photos for my review of {product}?",
        "I want to give 5 stars.",
        "How to share my feedback on the {product}?",
        "I want to review order {order_id}.",
        "Leave a comment on the product.",
        "How to add a product review?",
        "Can I rate the seller?"
    ],
    "warranty_info": [
        "What is the warranty on this {product}?",
        "Does it come with a guarantee?",
        "How many months of warranty for {product}?",
        "Is the warranty valid in {city}?",
        "What does the warranty cover?",
        "Tell me about the warranty terms.",
        "My {product} is under warranty.",
        "How to claim guarantee?",
        "Is physical damage covered under warranty?",
        "Warranty period please."
    ],
    "bulk_order": [
        "I want to place a bulk order.",
        "Do you provide a discount for {quantity} pieces of {product}?",
        "I need to order {product} in bulk.",
        "Can I buy wholesale?",
        "B2B ordering process?",
        "I want to order {quantity} of {product}.",
        "Corporate order for our office in {city}.",
        "How to get bulk pricing?",
        "We are looking to purchase large quantities.",
        "Bulk discount available?"
    ],
    "gift_wrapping": [
        "Do you offer gift wrapping?",
        "Can you wrap this {product} as a gift?",
        "I want gift packaging for order {order_id}.",
        "Is there an option to send {product} as a gift?",
        "Add a gift message.",
        "Gift wrapping cost?",
        "Can you conceal the price tag?",
        "Send this wrapped to {city}.",
        "I need this to be a present.",
        "Gift wrap this please."
    ],
    "human_escalate": [
        "I want to talk to a human.",
        "Connect me to an agent.",
        "Let me speak to customer care.",
        "Transfer me to a real person.",
        "I need to speak to the manager.",
        "Call me back.",
        "Put me through to support staff.",
        "I am tired of talking to a bot.",
        "Human intervention required.",
        "Can I get a customer service representative?"
    ]
}

def generate_dataset():
    dataset = []
    
    for intent, templates in TEMPLATES.items():
        count = 0
        
        # Variations generators
        prefix_variations = ["Please ", "Hi, ", "Can you tell me ", "I want to know ", "Urgent: ", "", ""]
        suffix_variations = [" thanks.", " please.", " now.", " ASAP.", "", ""]

        # Generate examples
        while count < 200:
            for template in random.sample(templates, len(templates)):
                # Fill placeholders
                filled = template.format(
                    product=random.choice(PRODUCTS),
                    size=random.choice(SIZES),
                    city=random.choice(CITIES),
                    date=random.choice(DATES),
                    order_id="ORD" + str(random.randint(1000, 9999)),
                    coupon_code="OFF" + str(random.randint(10, 50)),
                    account_id="USR" + str(random.randint(100, 999)),
                    quantity=random.randint(5, 100)
                )
                
                # Add variations randomly
                final_utterance = random.choice(prefix_variations) + filled + random.choice(suffix_variations)
                
                # Cleanup spaces and caps
                final_utterance = final_utterance.strip()
                final_utterance = final_utterance[0].upper() + final_utterance[1:] if final_utterance else ""
                
                dataset.append({"text": final_utterance, "intent": intent})
                count += 1
                if count >= 200:
                    break

    return dataset

if __name__ == "__main__":
    generated_data = generate_dataset()
    random.shuffle(generated_data)
    
    out_file = os.path.join(os.path.dirname(__file__), '../../data/raw/english_base_intents.csv')
    
    with open(out_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['text', 'intent'])
        writer.writeheader()
        writer.writerows(generated_data)
        
    print(f"Generated {len(generated_data)} english intent examples perfectly distributed across '{out_file}'")
