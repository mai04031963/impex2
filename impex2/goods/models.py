from django.db import models


# Create your models here.
class Good(models.Model):
    name = models.CharField(null=True, max_length=255, verbose_name='Товар')          # наименование товара (раздела)
    article = models.CharField(null=True, max_length=20, blank=True, verbose_name='Артикул') # артикул
    catalog_number = models.CharField(null=True, max_length=20, blank=True, verbose_name='Каталожный номер') # КАТАЛОЖНЫЙ НОМЕР
    description = models.TextField(null=True, blank=True, verbose_name='Описание')      # описание товара
    in_stock = models.DecimalField(default=0, db_index=True, decimal_places=2, max_digits=6, verbose_name='Кол-во на складе')  # количество на складе
    cat1 = models.BigIntegerField(default=0, blank=False, verbose_name='id раздела 1-ого уровня')   # номер раздела 1-ого уровня
    cat2 = models.BigIntegerField(default=0, blank=False, verbose_name='id раздела 2-ого уровня')   # номер раздела 2-ого уровня
    cat3 = models.BigIntegerField(default=0, blank=False, verbose_name='id раздела 3-ого уровня')   # номер раздела 3-его уровня
    is_good = models.BooleanField(default=True, blank=False, verbose_name='Товар/ раздел')# True - товар, False - раздел
    price = models.DecimalField(default=0, blank=True, decimal_places=2, max_digits=12, verbose_name='Цена') # цена
    supplier = models.CharField(null=True, max_length=4, blank=True, verbose_name='Поставшик')   # поставщик

    def __str__(self):
        return self.name

    class Meta:
        verbose_name_plural = "Товары"
        verbose_name = "Товар"