from django import forms
from django.contrib.auth.models import User
from django.contrib.auth.forms import UserCreationForm

# Registration form
class RegisterForm(UserCreationForm):
    email = forms.EmailField(required=True,widget=forms.EmailInput(attrs={
        'placeholder':'Enter The Email',
        'class':'email'}))
    password1 = forms.CharField(
        label="Password",
        widget=forms.PasswordInput(attrs={
            "class": "form-control",
            "placeholder": "Enter The Password",
        })
    )
    password2 = forms.CharField(
        label="Confirm The Password",
        widget=forms.PasswordInput(attrs={
            "class": "form-control",
            "placeholder": "Confirm password",
        })
    )

    class Meta:
        model = User
        fields = ["username", "email", "password1", "password2"]
        widgets = {
            "username": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Enter User Name",
                "maxlength": "20"
            })
        }


    def clean_email(self):
        email = self.cleaned_data.get("email")
        if User.objects.filter(email=email).exists():
            raise forms.ValidationError("This email is already taken.")
        return email
