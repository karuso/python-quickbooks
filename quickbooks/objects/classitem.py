from six import python_2_unicode_compatible
from .base import QuickbooksManagedObject, Ref


@python_2_unicode_compatible
class ClassItem(QuickbooksManagedObject):
    """
    QBO definition: Class objects provide a way to track different
    segments of the business so they're not tied to a particular client
    or project. For example, you can define classes to break down the
    income and expenses for each business segment. Classes are available
    to the entire transaction or to individual detail lines of a transaction.
    """

    # class_dict = {
    #     "CompanyAddr": Address,
    #     "CustomerCommunicationAddr": Address,
    #     "LegalAddr": Address,
    #     "PrimaryPhone": PhoneNumber,
    #     "Email": EmailAddress,
    #     "WebAddr": WebAddress
    # }

    qbo_object_name = "Class"

    def __init__(self):
        super(ClassItem, self).__init__()

        self.Id = None
        self.Name = ""
        self.Active = True
        self.SubItem = False
        self.FullyQualifiedName = ""   # Readonly


    def __str__(self):
        return f"[{self.Id}] {self.Name} ({self.FullyQualifiedName})"

    def to_ref(self):
        ref = Ref()

        ref.name = self.Name
        ref.type = self.qbo_object_name
        ref.value = self.Id

        return ref
